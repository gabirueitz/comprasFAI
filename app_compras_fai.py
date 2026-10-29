import os
import json
import threading
import time
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
from playwright.sync_api import sync_playwright

CONFIG_FILE = "config.json"

def ler_tabela_flexivel(caminho):
    """Lê arquivos .csv (com vários encodings) ou planilhas .xlsx/.xls."""
    ext = os.path.splitext(caminho)[1].lower()
    
    if ext in ['.xlsx', '.xls']:
        df = pd.read_excel(caminho)
        df = df.dropna(how='all').dropna(axis=1, how='all')
        if df.shape[1] >= 4:
            return df
        raise ValueError(f"A planilha Excel possui apenas {df.shape[1]} coluna(s). São necessárias pelo menos 4.")

    codificacoes = ['utf-8-sig', 'utf-8', 'latin1', 'cp1252', 'iso-8859-1']
    delimitadores = [';', ',', '\t']
    
    for cod in codificacoes:
        for sep in delimitadores:
            try:
                df = pd.read_csv(caminho, sep=sep, encoding=cod)
                df = df.dropna(how='all').dropna(axis=1, how='all')
                if df.shape[1] >= 4:
                    return df
            except Exception:
                continue

    for cod in ['latin1', 'cp1252', 'utf-8']:
        try:
            df = pd.read_csv(caminho, sep=None, engine='python', encoding=cod)
            df = df.dropna(how='all').dropna(axis=1, how='all')
            if df.shape[1] >= 4:
                return df
        except Exception:
            continue
            
    raise ValueError("Não foi possível identificar a codificação ou o separador do arquivo CSV/planilha.")

def calcular_data_recebimento():
    """Calcula a data atual + 15 dias corridos. Adia para segunda se coincidir com fim de semana."""
    hoje = datetime.now()
    data_alvo = hoje + timedelta(days=15)
    if data_alvo.weekday() == 5:
        data_alvo += timedelta(days=2)
    elif data_alvo.weekday() == 6:
        data_alvo += timedelta(days=1)
    return data_alvo.strftime("%d/%m/%Y")

class AplicativoComprasFAI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Automação de Compras - FAI UFSCar")
        self.geometry("660x650")
        self.resizable(False, False)

        self.config_data = self.carregar_config()
        self.criar_widgets()

    def carregar_config(self):
        padrao = {
            "nome": "", "email": "", "telefone": "",
            "cidade": "São Carlos", "endereco": "",
            "departamento": "", "codigo_alinea": "1900",
            "ultimo_projeto": "17172",
            "informacoes_complementares": ""
        }
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    padrao.update(json.load(f))
            except Exception:
                pass
        return padrao

    def salvar_config(self):
        for k in self.config_data:
            if hasattr(self, f"var_{k}"):
                self.config_data[k] = getattr(self, f"var_{k}").get()
        self.config_data["ultimo_projeto"] = self.var_projeto.get().strip()
        self.config_data["informacoes_complementares"] = self.txt_info_comp.get("1.0", tk.END).strip()
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(self.config_data, f, ensure_ascii=False, indent=2)

    def criar_widgets(self):
        notebook = ttk.Notebook(self)
        notebook.pack(fill=tk.BOTH, expand=True, padx=8, pady=6)

        # ----------------- Aba 1: Execução do Pedido -----------------
        aba_exec = ttk.Frame(notebook)
        notebook.add(aba_exec, text="Novo Pedido")

        # Linha 1: Projeto e Planilha
        frame_topo = ttk.Frame(aba_exec)
        frame_topo.pack(fill=tk.X, padx=8, pady=(4, 2))

        frame_proj = ttk.LabelFrame(frame_topo, text="Projeto de Destino")
        frame_proj.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 4))

        self.var_projeto = tk.StringVar(value=self.config_data.get("ultimo_projeto", "17172"))
        ttk.Label(frame_proj, text="Nº:").pack(side=tk.LEFT, padx=(5, 2), pady=3)
        entry_projeto = ttk.Entry(frame_proj, textvariable=self.var_projeto, width=12)
        entry_projeto.pack(side=tk.LEFT, padx=(0, 5), pady=3)

        frame_csv = ttk.LabelFrame(frame_topo, text="Planilha do Pedido (.csv ou .xlsx)")
        frame_csv.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(4, 0))

        self.var_caminho_csv = tk.StringVar()
        ttk.Entry(frame_csv, textvariable=self.var_caminho_csv).pack(side=tk.LEFT, padx=4, pady=3, fill=tk.X, expand=True)
        ttk.Button(frame_csv, text="Procurar...", command=self.selecionar_csv).pack(side=tk.RIGHT, padx=4, pady=3)

        # Linha 2: Campos de Texto do Pedido (Objetivo e Empresa / Fornecedor)
        frame_especifico = ttk.LabelFrame(aba_exec, text="Dados Específicos deste Pedido")
        frame_especifico.pack(fill=tk.X, padx=8, pady=3)

        ttk.Label(frame_especifico, text="Objetivo / Justificativa:").pack(anchor=tk.W, padx=6, pady=(3, 0))
        self.txt_justificativa = scrolledtext.ScrolledText(frame_especifico, height=3)
        self.txt_justificativa.pack(fill=tk.X, padx=6, pady=(1, 4))

        ttk.Label(frame_especifico, text="Empresa / Fornecedor (Nome, Razão Social, E-mail, Telefone):").pack(anchor=tk.W, padx=6, pady=(2, 0))
        self.txt_fornecedor = scrolledtext.ScrolledText(frame_especifico, height=3)
        self.txt_fornecedor.pack(fill=tk.X, padx=6, pady=(1, 4))

        # Botão Principal
        self.btn_iniciar = ttk.Button(aba_exec, text="Iniciar Automação no Edge", command=self.iniciar_thread)
        self.btn_iniciar.pack(fill=tk.X, padx=8, pady=4)

        # Log de status
        frame_log = ttk.LabelFrame(aba_exec, text="Progresso da Execução")
        frame_log.pack(fill=tk.BOTH, expand=True, padx=8, pady=(2, 6))

        self.txt_log = scrolledtext.ScrolledText(frame_log, state=tk.DISABLED, height=6)
        self.txt_log.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

        # ----------------- Aba 2: Configurações Fixas -----------------
        aba_config = ttk.Frame(notebook)
        notebook.add(aba_config, text="Configurações Fixas")

        frame_campos = ttk.LabelFrame(aba_config, text="Dados do Solicitante e Entrega")
        frame_campos.pack(fill=tk.X, padx=8, pady=4)

        campos = [
            ("Nome Solicitante:", "nome"),
            ("E-mail:", "email"),
            ("Telefone:", "telefone"),
            ("Cidade de Entrega:", "cidade"),
            ("Endereço de Entrega:", "endereco"),
            ("Departamento / Laboratório:", "departamento"),
            ("Código da Alínea Padrão:", "codigo_alinea")
        ]

        for rotulo, chave in campos:
            f = ttk.Frame(frame_campos)
            f.pack(fill=tk.X, padx=6, pady=2)
            ttk.Label(f, text=rotulo, width=24, anchor=tk.W).pack(side=tk.LEFT)
            var = tk.StringVar(value=self.config_data.get(chave, ""))
            setattr(self, f"var_{chave}", var)
            ttk.Entry(f, textvariable=var).pack(side=tk.RIGHT, fill=tk.X, expand=True)

        frame_comp = ttk.LabelFrame(aba_config, text="Informações Complementares (Texto Fixo)")
        frame_comp.pack(fill=tk.BOTH, expand=True, padx=8, pady=4)

        self.txt_info_comp = scrolledtext.ScrolledText(frame_comp, height=4)
        self.txt_info_comp.insert(tk.END, self.config_data.get("informacoes_complementares", ""))
        self.txt_info_comp.pack(fill=tk.BOTH, expand=True, padx=4, pady=4)

        ttk.Button(aba_config, text="Salvar Configurações", command=self.btn_salvar_config_click).pack(pady=6)

    def btn_salvar_config_click(self):
        self.salvar_config()
        messagebox.showinfo("Sucesso", "Configurações salvas com sucesso!")

    def selecionar_csv(self):
        caminho = filedialog.askopenfilename(
            filetypes=[
                ("Planilhas e CSV", "*.csv;*.xlsx;*.xls"),
                ("Arquivos CSV", "*.csv"),
                ("Arquivos Excel", "*.xlsx;*.xls"),
                ("Todos os arquivos", "*.*")
            ]
        )
        if caminho:
            self.var_caminho_csv.set(caminho)

    def log(self, mensagem):
        self.txt_log.config(state=tk.NORMAL)
        self.txt_log.insert(tk.END, mensagem + "\n")
        self.txt_log.see(tk.END)
        self.txt_log.config(state=tk.DISABLED)

    def iniciar_thread(self):
        caminho_csv = self.var_caminho_csv.get().strip()
        num_proj = self.var_projeto.get().strip()
        if not num_proj:
            messagebox.showerror("Erro", "Por favor, digite o número do projeto.")
            return
        if not caminho_csv or not os.path.exists(caminho_csv):
            messagebox.showerror("Erro", "Por favor, selecione uma planilha ou arquivo CSV válido.")
            return

        self.salvar_config()
        self.btn_iniciar.config(state=tk.DISABLED)
        thread = threading.Thread(target=self.executar_automacao, daemon=True)
        thread.start()

    def garantir_formulario_aberto(self, context, num_projeto):
        """
        Navega até o projeto e seleciona Projeto e 'SC - Solicitação de Compra'
        usando a seleção por jQuery e disparando o clique no botão '+' pelo DOM.
        """
        for p in context.pages:
            try:
                if p.locator("#NomeContato").is_visible():
                    self.log("Formulário de SC detectado como aberto!")
                    return p
            except Exception:
                pass

        page = context.pages[-1]

        # Localiza o card do projeto exato
        seletor_link_exato = f"a[href*='/Coordenadores/Formulario/Listar/{num_projeto}']"
        link_form = page.locator(seletor_link_exato).first

        if link_form.is_visible():
            self.log(f"Acessando página de formulários do projeto {num_projeto}...")
            link_form.click()
            page.wait_for_load_state("domcontentloaded")
            time.sleep(1.5)
        else:
            cards = page.locator("div.box, div.col-md-4, div.col-sm-6")
            qtd_cards = cards.count()
            for i in range(qtd_cards):
                card = cards.nth(i)
                txt_titulo = card.locator("h3, .box-title").first.inner_text()
                if num_projeto in txt_titulo:
                    self.log(f"Card localizado: {txt_titulo.strip()[:35]}...")
                    botao_form = card.locator("a, button, span").filter(has_text="Formulários").first
                    botao_form.click()
                    page.wait_for_load_state("domcontentloaded")
                    time.sleep(1.5)
                    break

        # Na barra 'Adicionar Novo Formulário': seleciona os menus e clica no botão (+) via DOM
        self.log(f"Configurando Projeto {num_projeto} e 'SC - Solicitação de Compra'...")

        page.evaluate(f"""() => {{
            $('select').each(function() {{
                const $sel = $(this);
                const opt = $sel.find('option').filter(function() {{ return $(this).text().indexOf('{num_projeto}') !== -1; }});
                if (opt.length) {{
                    $sel.val(opt.val()).trigger('change');
                }}
            }});

            $('select').each(function() {{
                const $sel = $(this);
                const opt = $sel.find('option').filter(function() {{ return $(this).text().indexOf('Solicitação de Compra') !== -1; }});
                if (opt.length) {{
                    $sel.val(opt.val()).trigger('change');
                }}
            }});

            setTimeout(() => {{
                const btn = document.querySelector('#adicionar') || document.querySelector('a.btn-success');
                if (btn) {{
                    btn.classList.remove('disabled');
                    btn.removeAttribute('disabled');
                    btn.click();
                }}
            }}, 300);
        }}""")

        self.log("Aguardando carregamento da página de cadastro...")

        inicio = time.time()
        while time.time() - inicio < 25:
            for p in context.pages:
                try:
                    if p.locator("#NomeContato").is_visible():
                        p.wait_for_load_state("domcontentloaded")
                        self.log("Formulário carregado com sucesso!")
                        return p
                except Exception:
                    pass
            time.sleep(1)

        messagebox.showinfo("Ação Necessária", "Abra a página do formulário no Edge e clique em OK para prosseguir.")
        return context.pages[-1]

    def executar_automacao(self):
        try:
            caminho_csv = self.var_caminho_csv.get().strip()
            num_projeto = self.var_projeto.get().strip()

            try:
                tabela = ler_tabela_flexivel(caminho_csv)
            except Exception as e:
                self.log(f"Erro ao carregar planilha: {e}")
                messagebox.showerror("Erro na Planilha", f"Não foi possível ler as colunas corretamente:\n{e}")
                return

            self.log(f"Planilha carregada: {len(tabela)} itens detectados.")

            justificativa = self.txt_justificativa.get("1.0", tk.END).strip()
            fornecedor = self.txt_fornecedor.get("1.0", tk.END).strip()
            data_limite = calcular_data_recebimento()

            self.log("Iniciando o Microsoft Edge...")
            user_data_dir = os.path.join(os.getcwd(), "sessao_edge_playwright")
            p = sync_playwright().start()

            context = p.chromium.launch_persistent_context(
                user_data_dir=user_data_dir,
                channel="msedge",
                headless=False,
                slow_mo=100,
                viewport=None,
                args=["--disable-blink-features=AutomationControlled", "--start-maximized"],
                ignore_default_args=["--enable-automation"]
            )

            page = context.pages[0] if context.pages else context.new_page()
            page.goto("https://sistemas.fai.ufscar.br/Coordenadores/Projeto/Listar", wait_until="domcontentloaded")

            self.log("--> Faça o login no Edge (se necessário).")
            messagebox.showinfo(
                "Ação Necessária",
                "1. Certifique-se de que está logado no sistema.\n"
                "2. Deixe na tela de Projetos ou já abra o formulário.\n"
                "3. Clique em OK para iniciar o preenchimento automático."
            )

            # Localiza a página do formulário
            page_form = self.garantir_formulario_aberto(context, num_projeto)

            # Preenchimento do cabeçalho
            self.log("Preenchendo cabeçalho e dados de contato...")

            page_form.fill("#NomeContato", self.config_data.get("nome", ""))
            page_form.fill("#EmailContato", self.config_data.get("email", ""))
            page_form.fill("#TelefoneContato", self.config_data.get("telefone", ""))
            page_form.fill("#LocalEntrega", self.config_data.get("endereco", ""))
            page_form.fill("#DepartamentoEntrega", self.config_data.get("departamento", ""))

            if self.config_data.get("cidade"):
                page_form.fill("#CidadeEntrega", self.config_data.get("cidade"))

            page_form.fill("#Vencimento", data_limite)
            page_form.evaluate(f"$('#Vencimento').val('{data_limite}').trigger('change');")

            if justificativa:
                page_form.fill("#justificativaDescricao", justificativa)
            if fornecedor:
                page_form.fill("#fornecedorDescricao", fornecedor)
            if self.config_data.get("informacoes_complementares"):
                page_form.fill("#observacoesDescricao", self.config_data.get("informacoes_complementares"))

            codigo_alinea = self.config_data.get("codigo_alinea", "1900")

            # Preenchimento dos itens da planilha
            for i, linha in tabela.iterrows():
                descricao = str(linha.iloc[0])
                quantidade = str(linha.iloc[1]).replace(".", ",")
                valor = str(linha.iloc[2]).replace(".", ",")
                unidade_csv = str(linha.iloc[3]).strip()

                self.log(f"Preenchendo item {i+1}: {descricao[:30]}...")

                if page_form.locator("#mensagemSemItens").is_visible():
                    page_form.locator("#mensagemSemItens").click()
                else:
                    page_form.locator("#addItemButton").click()

                modal = page_form.locator("#modal_item")
                modal.wait_for(state="visible", timeout=10000)

                page_form.fill("#DescItem", descricao)

                page_form.evaluate(f"$('#CtaFinanceiraItem').val('{codigo_alinea}').trigger('change');")
                page_form.evaluate("""() => {
                    const primeiroCC = $('#CentroCustoItem option:not(:first)').val();
                    if (primeiroCC) $('#CentroCustoItem').val(primeiroCC).trigger('change');
                }""")

                # Mapeamento de unidades
                page_form.evaluate("""(unidadeBusca) => {
                    function normalizar(str) {
                        return (str || "").normalize("NFD").replace(/[\\u0300-\\u036f]/g, "").toUpperCase().trim();
                    }
                    const termo = normalizar(unidadeBusca);
                    const sinonimos = {
                        "UNIDADE": ["UN", "UND", "UNID", "UNIDADE", "PECA", "PC"],
                        "UN": ["UNIDADE", "UND", "UNID", "PECA", "PC"],
                        "PACOTE": ["PCT", "PAC", "PACOTE", "CX", "CAIXA"],
                        "PCT": ["PACOTE", "PAC", "CX", "CAIXA"],
                        "PECA": ["PC", "UN", "UNIDADE"],
                        "METRO": ["M", "MT", "METRO"],
                        "ROLO": ["RL", "ROLO"],
                        "CONJUNTO": ["CJ", "CONJUNTO", "KIT"],
                        "KIT": ["CJ", "CONJUNTO", "KIT"]
                    };
                    let opcao = $('#UnMedidaItem option').filter(function() {
                        return normalizar($(this).text()) === termo;
                    });
                    if (!opcao.length) {
                        opcao = $('#UnMedidaItem option').filter(function() {
                            const txt = normalizar($(this).text());
                            return txt.includes(termo) || termo.includes(txt);
                        });
                    }
                    if (!opcao.length && sinonimos[termo]) {
                        const lista = sinonimos[termo];
                        opcao = $('#UnMedidaItem option').filter(function() {
                            return lista.includes(normalizar($(this).text()));
                        });
                    }
                    if (opcao.length && opcao.val()) {
                        $('#UnMedidaItem').val(opcao.val()).trigger('change');
                    }
                }""", unidade_csv)

                page_form.fill("#QtdItem", quantidade)
                page_form.fill("#VlrUnItem", valor)
                page_form.evaluate("() => { if (typeof CalculaNovo === 'function') CalculaNovo(); }")

                page_form.locator("#saveItemBtn").click()
                modal.wait_for(state="hidden", timeout=10000)

            self.log("\nProcesso finalizado com sucesso!")
            messagebox.showinfo("Concluído", "Todos os itens e dados foram inseridos com sucesso!\nRevise no navegador e salve o rascunho.")

        except Exception as e:
            self.log(f"\nERRO: {e}")
            messagebox.showerror("Erro de Execução", f"Ocorreu uma falha: {e}")
        finally:
            self.btn_iniciar.config(state=tk.NORMAL)

if __name__ == "__main__":
    app = AplicativoComprasFAI()
    app.mainloop()
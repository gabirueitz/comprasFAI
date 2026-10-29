# ComprasFAI - Automação de Solicitação de Compra (FAI UFSCar)

Aplicativo desktop com interface gráfica para automação do preenchimento de **Solicitações de Compra (SC)** no portal de Coordenadores da **FAI UFSCar**, a partir de planilhas Excel (`.xlsx`) ou arquivos `.csv`.

---

## Como Usar (Versão Executável)

1. Acesse a aba **[Releases](https://github.com/gabirueitz/comprasFAI/releases)** e baixe o arquivo `ComprasFAI_Windows.zip`.
2. Extraia o arquivo `.zip` em uma pasta de sua preferência.
3. Execute o aplicativo dando duplo clique em **`ComprasFAI.exe`**.
4. Na aba **Configurações Fixas**, preencha seus dados de solicitante, endereço de entrega e alínea padrão (essas informações ficam salvas localmente).
5. Na aba **Novo Pedido**, insira o número do projeto, selecione sua planilha e clique em **Iniciar Automação no Edge**.

---

## Formato Obrigatório da Planilha

O arquivo (`.xlsx` ou `.csv`) deve conter pelo menos **4 colunas**, exatamente nesta ordem:

| Coluna 1 (A)          | Coluna 2 (B)   | Coluna 3 (C)            | Coluna 4 (D)          |
| :-------------------- | :------------- | :---------------------- | :-------------------- |
| **Descrição do Item** | **Quantidade** | **Valor Unitário (R$)** | **Unidade de Medida** |
| Cabo de Rede Cat6 2m  | 5              | 18,50                   | UN                    |
| Reagente Químico PA   | 2              | 145,00                  | FR                    |

_O sistema aceita unidades comuns como `UN`, `PCT`, `CX`, `M`, `L`, etc._

---

## Rodando pelo Código-Fonte (Desenvolvedores)

### Pré-requisitos

- Python 3.10 ou superior
- Microsoft Edge instalado

```bash
# 1. Clonar o repositório
git clone [https://github.com/gabirueitz/comprasFAI.git](https://github.com/gabirueitz/comprasFAI.git)
cd comprasFAI

# 2. Instalar dependências
pip install -r requirements.txt

# 3. Instalar o navegador do Playwright
playwright install chromium

# 4. Executar
python app_compras_fai.py
```

# 🚀 Automação de Lançamento de Despesas - SSW Tela 475

Sistema ultraleve e rápido para preenchimento e gravação automática de despesas no sistema SSW a partir de dados da planilha Excel.

---

## 📋 Mapeamento Exato dos Campos (Sequência da Planilha)

Cada linha que você copiar e colar no app preencherá exatamente na ordem:

| Nº | Campo da Planilha | Tela do SSW | Elemento / ID no SSW | Exemplo |
|---|---|---|---|---|
| 1 | **Unidade** | Parte 1 | `<input id="3">` | `FSP` |
| 2 | **Fornecedor / CNPJ** | Parte 1 | `<input id="chave_nfe">` | `14821124000142` |
| 3 | **Evento** | Parte 1 | `<input id="5">` | `7022` ou `7564` |
| - | *Avançar para Parte 2* | Parte 1 | `<a id="6">►</a>` | *(Clique automático)* |
| 4 | **Série** | Parte 2 | `<input id="4">` | `1` |
| 5 | **Nº Compra** | Parte 2 | `<input id="5">` | `10426197` |
| 6 | **Modelo** | Parte 2 | `<input id="7">` | `98` |
| 7 | **Valor** | Parte 2 | `<input id="15">` | `13,20` |
| 8 | **Data de Emissão** | Parte 2 | `<input id="16">` | `170926` |
| 9 | **Data de Entrada** | Parte 2 | `<input id="17">` | `180926` |
| 10 | **Data de Vencimento** | Parte 2 | `<input id="data_vcto">` | `170926` |
| 11 | **Data de Pagamento** | Parte 2 | `<input id="data_pgto">` | `170926` |
| 12 | **Competência** | Parte 2 | `<input id="mes_competencia">` | `0926` |
| 13 | **Valor da Parcela** | Parte 2 | `<input id="vlr_parcela">` | `13,20` |
| 14 | **Histórico** | Parte 2 | `<input id="historico">` | `S/N - KMR9I57 - TAG...` |
| - | *Gravar lançamento* | Parte 2 | `<a id="lnk_grava_lancto">` | *(Gravação e retorno automático)* |

---

## ⚡ Como Usar (Passo a Passo Rápido)

### 1. Iniciar o Dashboard
- Dê dois cliques em **`iniciar_dashboard.bat`** (ou abra o arquivo `dashboard/index.html` diretamente no seu Chrome).
- O painel abrirá instantaneamente.

### 2. Copiar os Dados da Planilha
- Abra a sua planilha Excel (`BASE - TARGET.xlsx` ou qualquer outra).
- Selecione as linhas desejadas (todas as colunas de dados).
- Pressione **Ctrl + C**.

### 3. Colar no Dashboard
- Vá para a tela do Dashboard e pressione **Ctrl + V** (ou clique na área tracejada).
- Os dados serão importados imediatamente para a tabela, com as datas, valores e competências já convertidos e validados no formato exigido pelo SSW!
- Se preferir, clique no botão **`⚡ Carregar Planilha TARGET`** para carregar os registros da base padrão.

### 4. Ativar a Extensão no Chrome (Apenas na 1ª vez - 15 segundos)
1. No seu Chrome, abra `chrome://extensions/`
2. Ative a chave **"Modo do desenvolvedor"** no canto superior direito.
3. Clique no botão **"Carregar sem compactação"** e selecione a pasta:
   `c:\Users\Win\Desktop\automaçao target\extension`
4. Pronto! O script já estará ativo na aba do SSW.

### 5. Iniciar a Execução
1. Deixe o SSW aberto na Tela 475: `https://sistema.ssw.inf.br/bin/ssw0094`.
2. No Dashboard, clique em **`▶ Iniciar Automação`**.
3. O robô preencherá linha por linha, gravará cada lançamento, capturará o número do lançamento gerado (ex: `FSP236457`) e voltará automaticamente para a tela 475 para fazer a próxima linha!
4. Ao final, você pode clicar em **`📥 Exportar Log`** para salvar o relatório completo em CSV com os números gerados.

---

## 🎭 Opção Direta com Playwright (100% Robusto e Visível)

Se preferir rodar a automação profissional via **Playwright**:
1. Dê dois cliques no arquivo **`rodar_automacao_playwright.bat`**.
2. Uma janela visível do Chrome se abrirá.
3. Se pedir login, basta digitar sua senha e entrar.
4. O Playwright assume imediatamente o controle, preenche cada campo pelo seu ID exato (`#3`, `#chave_nfe`, `#5`, `#6`, `#4`, etc.), clica em Gravar e processa todas as linhas da planilha de ponta a ponta sem qualquer erro!


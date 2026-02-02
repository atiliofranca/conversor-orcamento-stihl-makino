# Orçamento Stihl - Conversor de Planilhas para CSV

Ferramenta simples para converter planilhas Excel de orçamento (com colunas/células “REFERÊNCIA” e “QTDE.”) em arquivos CSV prontos para importação, seguindo o padrão do arquivo de amostra `br_cart_import_format_csv_sample.csv`.

## Como usar

1. Baixe o executável `orcamento-stihl.exe` na seção [Releases](https://github.com/SEU_USUARIO/SEU_REPOSITORIO/releases) do GitHub.
2. Dê um duplo clique no arquivo para abrir o programa (não precisa instalar nada).
3. Na janela que abrir, selecione sua planilha Excel de orçamento.
4. O programa processa os dados automaticamente, localiza as colunas/células “REFERÊNCIA” e “QTDE.”, e gera um novo arquivo CSV.
5. Escolha onde salvar o arquivo gerado. O nome sugerido é `orçamento-stihl.csv`.

**Formato do CSV gerado:**
Cada linha segue o padrão: `REFERENCIA,QTDE,` (sem cabeçalho, vírgula final intencional).

# Especificações Técnicas e Guia para Desenvolvedores

Este documento contém informações técnicas e instruções para quem deseja modificar o código ou gerar um novo executável da aplicação Orçamento Stihl.

## Requisitos Técnicos

- Python 3.8 ou superior
- Windows (para gerar o .exe com PyInstaller)
- Dependências: `pandas`, `openpyxl`, `xlrd`, `pyinstaller` (listadas em `requirements.txt`)

## Como rodar o código em modo desenvolvimento

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python main.py
```

## Como gerar o executável (.exe)

- Manualmente:

  ```powershell
  pip install pyinstaller
  pyinstaller --onefile --windowed --name orcamento-stihl main.py
  ```

  O executável será gerado em `dist\orcamento-stihl.exe`.
- Automaticamente:
  Execute o script `build_exe.bat` na raiz do projeto:

  ```bat
  build_exe.bat
  ```

  O executável será copiado para a pasta `dist\orcamento-stihl.exe`.

## Sobre o funcionamento

- O programa detecta automaticamente as colunas/células “REFERÊNCIA” e “QTDE.”, mesmo que estejam fora do cabeçalho padrão.
- Remove traços e caracteres não numéricos da referência.
- Converte a quantidade para inteiro, removendo casas decimais.
- Gera o CSV compatível com o sistema de importação Stihl.

## Observações técnicas

- Arquivos `.xls` exigem `xlrd` instalado no ambiente de build/execução; o código seleciona o engine apropriado pela extensão do arquivo.
- O leitor tenta reconhecer os títulos “REFERÊNCIA” e “QTDE” mesmo quando não estão no cabeçalho padrão, buscando as palavras nas células da planilha.
- O CSV de saída não contém cabeçalho e segue o padrão de amostra (vírgula final em cada linha) para compatibilidade com o sistema de destino.

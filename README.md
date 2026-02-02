# Orçamento Stihl - Gerador de CSV

Aplicativo simples para selecionar uma planilha Excel, extrair as colunas `REFERÊNCIA` e `QTDE.` e gerar um arquivo CSV no formato de `br_cart_import_format_csv_sample.csv`.

Características:
- Remove caracteres não numéricos da coluna `REFERÊNCIA` (retém apenas dígitos).
- Converte `QTDE.` para número inteiro (remove casas decimais, ex: `6,00` → `6`).
- Gera CSV sem cabeçalho, cada linha no formato `REFERENCIA,QTDE,` (note a vírgula final para coincidir com o arquivo de amostra).
- Nome sugerido do arquivo: `orçamento-stihl.csv`, usuário pode escolher pasta e nome final.

Requisitos:
- Python 3.8+
- Instalar dependências:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Executar (modo de desenvolvimento):

```powershell
python main.py
```

Gerar .exe (Windows) com PyInstaller (manual):

```powershell
pip install pyinstaller
pyinstaller --onefile --windowed main.py

# O executável ficará em `dist\main.exe`. Renomeie para algo como `orcamento-stihl.exe` se desejar.
```

Construir o `.exe` automaticamente com o script `.bat` (Windows):

- O repositório contém `build_exe.bat` que automatiza a criação do executável. O script:
  - cria/usa um virtualenv em `.venv`;
  - ativa o venv e atualiza o `pip`;
  - instala as dependências listadas em `requirements.txt`;
  - executa o `pyinstaller --onefile --windowed --name orcamento-stihl main.py`;
  - copia o executável final para a pasta `release\orcamento-stihl.exe`.

Como usar (`cmd.exe`):

```bat
build_exe.bat
```

Observações sobre o `build_exe.bat`:
- O script exige que o `python` esteja disponível no PATH.
- O nome do executável final usa `orcamento-stihl.exe` (ASCII) para evitar problemas com codificação de nomes em builds.
- Se preferir, você pode executar os passos manualmente (ver seção anterior).

Uso:
- Ao executar o programa, uma janela solicitará que você selecione a planilha Excel.
- Depois do processamento, será aberta uma janela para escolher onde salvar o CSV. O nome sugerido é `orçamento-stihl.csv`.

Observações:
- O script tenta localizar as colunas mesmo que o nome tenha pequenas variações (por exemplo, sem acento ou com/sem ponto em `QTDE.`).
- Caso sua planilha contenha formatos numéricos regionais (vírgula como separador decimal), o script faz uma tentativa de conversão segura.
# Orçamento Stihl - Gerador de CSV

Aplicativo simples para selecionar uma planilha Excel, extrair as colunas `REFERÊNCIA` e `QTDE.` e gerar um arquivo CSV no formato de `br_cart_import_format_csv_sample.csv`.

Características:
- Remove caracteres não numéricos da coluna `REFERÊNCIA` (retém apenas dígitos).
- Converte `QTDE.` para número inteiro (remove casas decimais, ex: `6,00` → `6`).
- Gera CSV sem cabeçalho, cada linha no formato `REFERENCIA,QTDE,` (note a vírgula final para coincidir com o arquivo de amostra).
- Nome sugerido do arquivo: `orçamento-stihl.csv`, usuário pode escolher pasta e nome final.

Requisitos:
- Python 3.8+
- Instalar dependências:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Executar (modo de desenvolvimento):

```powershell
python main.py
```

Gerar .exe (Windows) com PyInstaller:

```powershell
pip install pyinstaller
pyinstaller --onefile --windowed main.py

# O executável ficará em `dist\main.exe`. Renomeie para algo como `orcamento-stihl.exe` se desejar.
```

Uso:
- Ao executar o programa, uma janela solicitará que você selecione a planilha Excel.
- Depois do processamento, será aberta uma janela para escolher onde salvar o CSV. O nome sugerido é `orçamento-stihl.csv`.

Observações:
- O script tenta localizar as colunas mesmo que o nome tenha pequenas variações (por exemplo, sem acento ou com/sem ponto em `QTDE.`).
- Caso sua planilha contenha formatos numéricos regionais (vírgula como separador decimal), o script faz uma tentativa de conversão segura.

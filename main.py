import re
import sys
import os
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd


def normalize_col_name(s: str) -> str:
    return re.sub(r"[^a-z0-9]", "", s.lower())


def find_column(df: pd.DataFrame, candidates):
    cols = {normalize_col_name(c): c for c in df.columns}
    for cand in candidates:
        key = normalize_col_name(cand)
        if key in cols:
            return cols[key]
    return None


def ask_user_for_columns(df: pd.DataFrame):
    cols = [str(c) for c in df.columns]

    sel = {"ref": None, "qty": None}

    def on_ok():
        sel["ref"] = var_ref.get()
        sel["qty"] = var_qty.get()
        top.destroy()

    def on_cancel():
        top.destroy()

    top = tk.Toplevel()
    top.title("Selecionar colunas")
    tk.Label(top, text="Selecione a coluna que contém a REFERÊNCIA:").pack(padx=10, pady=(10, 0))
    var_ref = tk.StringVar(value=cols[0] if cols else "")
    tk.OptionMenu(top, var_ref, *cols).pack(fill="x", padx=10)

    tk.Label(top, text="Selecione a coluna que contém a QTDE:").pack(padx=10, pady=(10, 0))
    var_qty = tk.StringVar(value=cols[1] if len(cols) > 1 else (cols[0] if cols else ""))
    tk.OptionMenu(top, var_qty, *cols).pack(fill="x", padx=10)

    btn_frame = tk.Frame(top)
    btn_frame.pack(pady=10)
    tk.Button(btn_frame, text="OK", width=10, command=on_ok).pack(side="left", padx=5)
    tk.Button(btn_frame, text="Cancelar", width=10, command=on_cancel).pack(side="left", padx=5)

    top.transient()
    top.grab_set()
    top.wait_window()

    return sel["ref"], sel["qty"]


def process_file(in_path: str):
    # read excel (xls or xlsx) selecting engine by extension
    ext = os.path.splitext(in_path)[1].lower()
    df = None
    if ext in (".xls",):
        # .xls requires xlrd >= 2.0.1
        try:
            import xlrd  # type: ignore
        except Exception:
            messagebox.showerror("Dependência ausente", "Arquivo .xls detectado mas a dependência opcional 'xlrd' não está instalada.\nInstale com: pip install xlrd")
            raise
        df = pd.read_excel(in_path, dtype=str, engine="xlrd")
    else:
        # prefer openpyxl for xlsx/xlsm, else let pandas choose
        if ext in (".xlsx", ".xlsm"):
            try:
                import openpyxl  # type: ignore
                df = pd.read_excel(in_path, dtype=str, engine="openpyxl")
            except Exception:
                # fallback to pandas default
                df = pd.read_excel(in_path, dtype=str)
        else:
            df = pd.read_excel(in_path, dtype=str)

    ref_col = find_column(df, ["REFERÊNCIA", "REFERENCIA", "REF"])
    qty_col = find_column(df, ["QTDE.", "QTDE", "QUANTIDADE", "QTD"])

    if ref_col is None or qty_col is None:
        # Try to find the keywords inside the sheet cells (not only headers)
        # read raw sheet without header
        ext = os.path.splitext(in_path)[1].lower()
        if ext in (".xls",):
            engine = "xlrd"
        elif ext in (".xlsx", ".xlsm"):
            engine = "openpyxl"
        else:
            engine = None

        if engine:
            try:
                if engine == "xlrd":
                    import xlrd  # type: ignore
                else:
                    import openpyxl  # type: ignore
            except Exception:
                # missing optional engine
                messagebox.showerror("Dependência ausente", f"Arquivo requer engine '{engine}' não instalada. Instale com: pip install {engine}")
                raise

        if engine:
            sheet_raw = pd.read_excel(in_path, header=None, dtype=str, engine=engine)
        else:
            sheet_raw = pd.read_excel(in_path, header=None, dtype=str)

        def search_keyword_in_cells(sheet, candidates):
            for r in range(sheet.shape[0]):
                for c in range(sheet.shape[1]):
                    val = sheet.iat[r, c]
                    if pd.isna(val):
                        continue
                    s = normalize_col_name(str(val))
                    for cand in candidates:
                        if normalize_col_name(cand) in s:
                            return r, c
            return None, None

        if ref_col is None:
            r_ref, c_ref = search_keyword_in_cells(sheet_raw, ["REFERÊNCIA", "REFERENCIA", "REF"])
        else:
            r_ref, c_ref = None, None

        if qty_col is None:
            r_qty, c_qty = search_keyword_in_cells(sheet_raw, ["QTDE.", "QTDE", "QUANTIDADE", "QTD"])
        else:
            r_qty, c_qty = None, None

        if (ref_col is None and (r_ref is None)) or (qty_col is None and (r_qty is None)):
            messagebox.showerror("Colunas não encontradas", "Não foi possível localizar as colunas 'REFERÊNCIA' e 'QTDE.' na planilha selecionada.")
            return None

        # Build series from raw sheet if we found header cells in-sheet
        if ref_col is None and r_ref is not None:
            ref_series = sheet_raw.iloc[r_ref + 1 :, c_ref].reset_index(drop=True)
        else:
            ref_series = df[ref_col].astype(str).reset_index(drop=True)

        if qty_col is None and r_qty is not None:
            qty_series = sheet_raw.iloc[r_qty + 1 :, c_qty].reset_index(drop=True)
        else:
            qty_series = df[qty_col].astype(str).reset_index(drop=True)

        # iterate over paired series
        out_rows = []
        maxlen = max(len(ref_series), len(qty_series))
        for i in range(maxlen):
            raw_ref = ref_series.iat[i] if i < len(ref_series) else ""
            raw_qty = qty_series.iat[i] if i < len(qty_series) else ""

            if pd.isna(raw_ref) or str(raw_ref).strip() == "":
                continue

            ref = re.sub(r"\D", "", str(raw_ref))
            if ref == "":
                continue

            qty_str = str(raw_qty).strip()
            if qty_str.lower() in ("nan", "none", ""):
                qty = "0"
            else:
                qty_str = qty_str.replace(".", "") if qty_str.count(",") == 1 and qty_str.count(".") > 0 and qty_str.find(",") > qty_str.find(".") else qty_str
                qty_str = qty_str.replace(",", ".")
                try:
                    qty_float = float(qty_str)
                    qty = str(int(qty_float))
                except Exception:
                    m = re.match(r"(\d+)", qty_str)
                    qty = m.group(1) if m else "0"

            out_rows.append((ref, qty))

        return out_rows

    out_rows = []
    for _, row in df.iterrows():
        raw_ref = row.get(ref_col, "")
        raw_qty = row.get(qty_col, "")

        if pd.isna(raw_ref):
            continue

        # Clean reference: keep only digits
        ref = re.sub(r"\D", "", str(raw_ref))
        if ref == "":
            continue

        # Clean quantity: remove thousand separators, handle comma decimal
        qty_str = str(raw_qty).strip()
        if qty_str.lower() in ("nan", "none", ""):
            qty = "0"
        else:
            # Replace comma decimal separator with dot
            qty_str = qty_str.replace(".", "") if qty_str.count(",") == 1 and qty_str.count(".") > 0 and qty_str.find(",") > qty_str.find(".") else qty_str
            qty_str = qty_str.replace(",", ".")
            try:
                qty_float = float(qty_str)
                qty = str(int(qty_float))
            except Exception:
                # fallback: extract integer part
                m = re.match(r"(\d+)", qty_str)
                qty = m.group(1) if m else "0"

        out_rows.append((ref, qty))

    return out_rows


def write_csv(out_rows, out_path: str):
    # Match sample format: REF,QTDE,  (trailing comma)
    with open(out_path, "w", encoding="utf-8-sig", newline="") as f:
        for ref, qty in out_rows:
            f.write(f"{ref},{qty},\n")


def main():
    root = tk.Tk()
    root.withdraw()

    messagebox.showinfo("Selecionar planilha", "Selecione a planilha de orçamento do autcom")
    file_path = filedialog.askopenfilename(title="Selecione a planilha Excel", filetypes=[("Excel files", "*.xls *.xlsx *.xlsm"), ("All files", "*")])
    if not file_path:
        return

    try:
        rows = process_file(file_path)
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao processar o arquivo:\n{e}")
        return

    if not rows:
        messagebox.showwarning("Sem dados", "Nenhum registro processável foi encontrado na planilha.")
        return

    # Suggest default filename
    initial_name = "orçamento-stihl.csv"
    save_path = filedialog.asksaveasfilename(defaultextension=".csv", initialfile=initial_name, filetypes=[("CSV files", "*.csv"), ("All files", "*")], title="Salvar CSV gerado")
    if not save_path:
        return

    try:
        write_csv(rows, save_path)
        messagebox.showinfo("Concluído", f"Arquivo salvo em:\n{save_path}")
    except Exception as e:
        messagebox.showerror("Erro ao salvar", f"Não foi possível salvar o arquivo:\n{e}")


if __name__ == "__main__":
    main()

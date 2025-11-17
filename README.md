import re
import sys
import time
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
from typing import Dict, List, Tuple, Optional, Union
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment
 
# ---------------- CONFIGURAÇÕES ----------------
BALANCETE_XLSX = r"Z:\MER\Contratados\Estagiários\Felipe Nascimento\Laura\Projeto Copilot\BalanceteDiárioPadrão_Teste.xlsx"  # <--- Modificar o caminho
DEM_PL_IN  = r"Z:\MER\Contratados\Estagiários\Felipe Nascimento\Laura\Projeto Copilot\Dem_PL_Modelo_preenchido.xlsx" # <--- Modificar o caminho
DEM_PL_OUT = r"Dem_PL_Modelo_preenchido.xlsx" # <--- Modificar o caminho
 
 
BALANCETE_SHEET: Optional[Union[str, int]] = None
COL_CONTA = "V"
COL_SALDO = "K"
SAFE_SAVE_WITH_SUFFIX = True
# Somatórios por bloco (item 7)
CEL_BLOCO_ACOES      = "J34"
CEL_BLOCO_RENDA_FIXA = "J40"
CEL_BLOCO_RECEITAS   = "J45"
CEL_BLOCO_DESPESAS   = "J55"
TOTAL_CELLS = {CEL_BLOCO_ACOES, CEL_BLOCO_RENDA_FIXA, CEL_BLOCO_RECEITAS, CEL_BLOCO_DESPESAS}
BLOCOS_RECONHECIDOS = {
    "Ações e Opções": "ACOES",
    "Renda fixa e outros valores mobiliários": "RENDA_FIXA",
    "Demais receitas": "RECEITAS",
    "Demais despesas": "DESPESAS",
}
# ✅ Formatação desejada:
# - separador de milhar
# - sem casas decimais
# - negativos entre parênteses
# - zero vira "-"
NUM_FMT_INT_MIL = "#,##0;(#,##0);-"
ALIGN_RIGHT = Alignment(horizontal="right")
# ---------------- FUNÇÕES ----------------
def excel_col_to_zero_based(col_letter: str) -> int:
    col_letter = col_letter.strip().upper()
    num = 0
    for ch in col_letter:
        if not ('A' <= ch <= 'Z'):
            raise ValueError(f"Coluna inválida: {col_letter}")
        num = num * 26 + (ord(ch) - ord('A') + 1)
    return num - 1
def _read_balancete_df(balancete_path: Path, sheet: Optional[Union[str, int]]) -> pd.DataFrame:
    result = pd.read_excel(balancete_path, sheet_name=0 if sheet is None else sheet)
    if isinstance(result, dict):
        result = next(iter(result.values()))
    return result
def build_account_map(balancete_path: Path, sheet, col_conta, col_saldo) -> Dict[str, float]:
    df = _read_balancete_df(balancete_path, sheet)
    idx_conta = excel_col_to_zero_based(col_conta)
    idx_saldo = excel_col_to_zero_based(col_saldo)
    s_conta = df.iloc[:, idx_conta].astype(str)
    s_saldo = pd.to_numeric(df.iloc[:, idx_saldo], errors="coerce").fillna(0.0)
    contas = s_conta.str.extract(r"(\d+)", expand=False)
    tmp = pd.DataFrame({"conta": contas, "saldo": s_saldo}).dropna(subset=["conta"])
    return tmp.groupby("conta")["saldo"].sum().to_dict()
def normalize_text_for_accounts(s: str) -> str:
    s = s.replace("\xa0", " ")
    s = s.replace("R$", "").replace(".", "")
    s = s.replace("—", "+").replace("–", "+").replace("-", "+")
    s = s.replace(";", "+").replace(",", "+")
    s = s.replace("’", "'").replace("`", "'")
    s = re.sub(r"\s+", "", s)
    return s
def parse_accounts_from_cell(val) -> List[str]:
    if val is None:
        return []
    s = normalize_text_for_accounts(str(val)).replace('"', "").replace("'", "")
    return re.findall(r"\d+", s)
def should_replace_cell(val) -> bool:
    if val is None:
        return False
    if isinstance(val, (int, float)):
        return True
    return bool(re.search(r"\d", normalize_text_for_accounts(str(val))))
def safe_save_workbook(wb, path_out: Path):
    try:
        wb.save(path_out)
        return path_out
    except PermissionError:
        ts = time.strftime("%Y%m%d_%H%M%S")
        alt = path_out.with_name(f"{path_out.stem}_{ts}{path_out.suffix}")
        wb.save(alt)
        return alt
# ✅ Item 6 – arredonda por mil
def round_thousands_cell(value_reais: float) -> int:
    return int((Decimal(value_reais) / Decimal(1000)).quantize(Decimal("0"), rounding=ROUND_HALF_UP))
def apply_int_mil_format(cell):
    cell.number_format = "General"
    cell.number_format = NUM_FMT_INT_MIL
    cell.alignment = ALIGN_RIGHT
 
 
from pathlib import Path
from typing import Dict
from openpyxl import load_workbook

def replace_in_dem_pl(dem_in: Path, dem_out: Path, acc_map: Dict[str, float]) -> Path:
    wb = load_workbook(dem_in, data_only=False)
    changes = []
    totals_por_conta = {}
    missing_codes = {}
    soma_blocos = {"ACOES": 0, "RENDA_FIXA": 0, "RECEITAS": 0, "DESPESAS": 0}
    bloco_atual = None

    for ws in wb.worksheets:
        for row in ws.iter_rows():
            col_a_val = row[0].value
            if isinstance(col_a_val, str):
                key = col_a_val.strip()
                if key in BLOCOS_RECONHECIDOS:
                    bloco_atual = BLOCOS_RECONHECIDOS[key]

            for cell in row:
                if cell.coordinate.upper() in TOTAL_CELLS:
                    continue
                if not should_replace_cell(cell.value):
                    continue

                raw_expr = str(cell.value)
                contas = parse_accounts_from_cell(raw_expr)
                if not contas:
                    continue

                total_reais = 0.0
                for c in contas:
                    v = float(acc_map.get(c, 0.0))
                    total_reais += v
                    totals_por_conta[c] = totals_por_conta.get(c, 0.0) + v
                    if c not in acc_map:
                        missing_codes[c] = missing_codes.get(c, 0) + 1

                val_mil = round_thousands_cell(total_reais)
                cell.value = val_mil
                apply_int_mil_format(cell)
                changes.append((f"{ws.title}!{cell.coordinate}", raw_expr, total_reais, val_mil))

                if bloco_atual in soma_blocos:
                    soma_blocos[bloco_atual] += val_mil

    # ---- ATUALIZAÇÕES NA 1ª PLANILHA ----
    ws0 = wb.worksheets[0]
    for coord, key in [
        (CEL_BLOCO_ACOES, "ACOES"),
        (CEL_BLOCO_RENDA_FIXA, "RENDA_FIXA"),
        (CEL_BLOCO_RECEITAS, "RECEITAS"),
        (CEL_BLOCO_DESPESAS, "DESPESAS"),
    ]:
        ws0[coord].value = soma_blocos[key]
        apply_int_mil_format(ws0[coord])

    CEL_TOTAL_GERAL = "J58"
    total_geral = (
        soma_blocos["ACOES"]
        + soma_blocos["RENDA_FIXA"]
        + soma_blocos["RECEITAS"]
        + soma_blocos["DESPESAS"]
    )
    ws0[CEL_TOTAL_GERAL].value = total_geral
    apply_int_mil_format(ws0[CEL_TOTAL_GERAL])

    # ---- SALVAR E RETORNAR CAMINHO EFETIVO ----
    out_path = Path(dem_out)
    if SAFE_SAVE_WITH_SUFFIX:
        out_path = safe_save_workbook(wb, out_path)
    else:
        wb.save(out_path)
 
 
 
def main():
    bal = Path(BALANCETE_XLSX)
    dem_in = Path(DEM_PL_IN)
    
    if not bal.exists():
        print("ERRO — Balancete não encontrado:", bal)
        sys.exit(1)
        
    if not dem_in.exists():
        print("ERRO — Modelo não encontrado:", dem_in)
        sys.exit(1)
        
    acc_map = build_account_map(bal, BALANCETE_SHEET, COL_CONTA, COL_SALDO)
    out_file = replace_in_dem_pl(dem_in, Path(DEM_PL_OUT), acc_map)
    print("\n[ OK ] Concluído!")
    print(f"Arquivo gerado: {out_file}")
    
if __name__ == "__main__":
    main()# Projeto-Eventos_DEM_PL

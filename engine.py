from __future__ import annotations
import os
import json
import re
import warnings
from datetime import datetime
from pathlib import Path
import pandas as pd

from channel_report_generator import (
    AXIO_FILE as DEFAULT_AXIO_FILE,
    OUTPUT_FILE as CHANNEL_REPORT_FILE,
    RETAIL_FILE as DEFAULT_RETAIL_FILE,
    generate_channel_report,
)
from workbook_styles import style_service_workbook

# Constants
BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = Path("/tmp") if os.getenv("VERCEL") else BASE_DIR

DEFAULT_SERVICE_FILE = BASE_DIR / "mi_smart_report (6).csv"
DEFAULT_SERVICE_MASTER_FILE = BASE_DIR / "current_service_master.xlsx"
FINAL_REPORT_FILE = OUTPUT_DIR / "final_report.xlsx"
ZONAL_REPORT_FILE = OUTPUT_DIR / "zonal_report.xlsx"

APP_NAME = "Xiaomi Daily Report Engine"
EXCEL_CONTENT_TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
CHANNEL_MASTER_PATTERNS = ("Master*.xlsb", "Master*.xlsx")
CHANNEL_MASTER_LOOKUP_LABEL = "Master*.xlsb or Master*.xlsx"

REQUIRED_SERVICE_COLUMNS = {"PAYMENT STATUS", "ASC Code", "CUSTOMER PRICE"}
NEW_SERVICE_MASTER_COLUMNS = {"Agency_Code", "Agency_Name", "Region"}
LEGACY_SERVICE_MASTER_COLUMNS = {"ASC_Code", "ASC_Name_BI", "Zone", "State"}

MONTH_NAME_TO_NUMBER = {
    "jan": 1,
    "january": 1,
    "feb": 2,
    "february": 2,
    "mar": 3,
    "march": 3,
    "apr": 4,
    "april": 4,
    "may": 5,
    "jun": 6,
    "june": 6,
    "jul": 7,
    "july": 7,
    "aug": 8,
    "august": 8,
    "sep": 9,
    "sept": 9,
    "september": 9,
    "oct": 10,
    "october": 10,
    "nov": 11,
    "november": 11,
    "dec": 12,
    "december": 12,
}
CHANNEL_MASTER_NAME_PATTERN = re.compile(
    r"^master\s+(?P<month>[a-z]+)'?(?P<year>\d{2,4})(?:\s*\((?P<version>\d+)\))?\.(?:xlsb|xlsx)$",
    re.IGNORECASE,
)


def parse_channel_master_name(path: Path) -> tuple[int, int, int] | None:
    match = CHANNEL_MASTER_NAME_PATTERN.match(path.name)
    if not match:
        return None

    month_name = match.group("month").lower()
    month_number = MONTH_NAME_TO_NUMBER.get(month_name)
    if month_number is None:
        return None

    year = int(match.group("year"))
    if year < 100:
        year += 2000
    version = int(match.group("version") or 0)
    return year, month_number, version


def channel_master_sort_key(path: Path) -> tuple[int, int, int, int, int, str]:
    parsed = parse_channel_master_name(path)
    suffix_rank = 0 if path.suffix.lower() == ".xlsb" else 1
    if parsed is not None:
        year, month, version = parsed
        return (0, -year, -month, -version, suffix_rank, path.name.lower())

    return (1, 0, 0, 0, suffix_rank, path.name.lower())


def resolve_channel_master_file(base_dir: Path = BASE_DIR) -> Path | None:
    override = os.getenv("CHANNEL_MASTER_FILE", "").strip()
    if override:
        override_path = Path(override)
        return override_path if override_path.is_absolute() else base_dir / override_path

    candidates: list[Path] = []
    for pattern in CHANNEL_MASTER_PATTERNS:
        candidates.extend(path for path in base_dir.glob(pattern) if path.is_file())

    if not candidates:
        return None
    return sorted(candidates, key=channel_master_sort_key)[0]


DEFAULT_CHANNEL_MASTER_FILE = resolve_channel_master_file() or (BASE_DIR / "Master.xlsb")

def normalise_truthy(value: object) -> bool:
    if isinstance(value, bool): return value
    return str(value).strip().upper() in {"TRUE", "1", "YES", "Y", "PAID"}

def clean_dimension(series: pd.Series) -> pd.Series:
    cleaned = series.fillna("Blank").astype(str).str.strip()
    return cleaned.mask(cleaned.eq("") | cleaned.str.lower().eq("nan"), "Blank")

def read_master_workbook(path: Path, sheet_name: int | str = 0) -> pd.DataFrame:
    suffix = path.suffix.lower()
    engine = "pyxlsb" if suffix == ".xlsb" else None
    return pd.read_excel(path, sheet_name=sheet_name, engine=engine)

def validate_columns(frame: pd.DataFrame, required: set[str], label: str) -> None:
    missing = sorted(required.difference(frame.columns))
    if missing:
        raise ValueError(f"{label} is missing required column(s): {', '.join(missing)}")

def ordered_with_blank_last(values: pd.Series) -> list[str]:
    unique_values = list(dict.fromkeys(values.tolist()))
    ordered = [value for value in unique_values if value != "Blank"]
    if "Blank" in unique_values: ordered.append("Blank")
    return ordered

def normalise_numeric_cell(value: object) -> int:
    if value in (None, "") or pd.isna(value): return 0
    return int(round(float(value)))

def fill_empty_numeric_cells(worksheet: object, *, numeric_columns: tuple[int, ...], min_row: int) -> None:
    for row in worksheet.iter_rows(min_row=min_row):
        for column in numeric_columns:
            cell = row[column - 1]
            cell.value = normalise_numeric_cell(cell.value)

def normalise_service_master(master: pd.DataFrame) -> tuple[pd.DataFrame, str]:
    if NEW_SERVICE_MASTER_COLUMNS.issubset(master.columns):
        state_col = "State-2" if "State-2" in master.columns else "State"
        return pd.DataFrame({
            "ServiceCode": master["Agency_Code"],
            "Region": master["Region"],
            "State": master.get(state_col, "Unknown"),
            "ServiceCenter": master["Agency_Name"],
        }), "Agency_Code"
    if LEGACY_SERVICE_MASTER_COLUMNS.issubset(master.columns):
        return pd.DataFrame({
            "ServiceCode": master["ASC_Code"],
            "Region": master["Zone"],
            "State": master["State"],
            "ServiceCenter": master["ASC_Name_BI"],
        }), "ASC_Code"
    raise ValueError("Master workbook missing required service master columns.")

def build_final_rows(report_df: pd.DataFrame) -> list[dict[str, object]]:
    final_rows, g_unit, g_gwp = [], 0, 0.0
    for region in ordered_with_blank_last(report_df["Region"]):
        region_df = report_df[report_df["Region"] == region]
        r_unit, r_gwp, first_region = 0, 0.0, True
        for state in ordered_with_blank_last(region_df["State"]):
            state_df = region_df[region_df["State"] == state]
            s_unit, s_gwp = int(state_df["Unit"].sum()), float(state_df["GWP"].sum())
            r_unit += s_unit; r_gwp += s_gwp; g_unit += s_unit; g_gwp += s_gwp
            first_state = True
            for _, row in state_df.iterrows():
                final_rows.append({
                    "Region": region if first_region else "",
                    "State": state if first_state else "",
                    "Service Center Name": row["ServiceCenter"],
                    "Unit": int(row["Unit"]),
                    "GWP": int(round(float(row["GWP"]))),
                })
                first_region = first_state = False
            final_rows.append({"Region": "", "State": f"{state} Total", "Service Center Name": "", "Unit": s_unit, "GWP": int(round(s_gwp))})
        final_rows.append({"Region": f"{region} Total", "State": "", "Service Center Name": "", "Unit": r_unit, "GWP": int(round(r_gwp))})
    final_rows.append({"Region": "Grand Total", "State": "", "Service Center Name": "", "Unit": g_unit, "GWP": int(round(g_gwp))})
    return final_rows

def generate_service_report(service_path: Path, master_path: Path) -> dict[str, object]:
    service = pd.read_csv(service_path)
    validate_columns(service, REQUIRED_SERVICE_COLUMNS, "Service report")
    master = read_master_workbook(master_path)
    try:
        master, code_col = normalise_service_master(master)
    except:
        master = read_master_workbook(master_path, sheet_name=1)
        master, code_col = normalise_service_master(master)
    
    input_rows = len(service)
    service["ASC Code"] = service["ASC Code"].astype(str).str.strip()
    service["CUSTOMER PRICE"] = pd.to_numeric(service["CUSTOMER PRICE"], errors="coerce").fillna(0)
    service = service[service["PAYMENT STATUS"].map(normalise_truthy)].copy()
    
    merged = service.merge(master, left_on="ASC Code", right_on="ServiceCode", how="left")
    unmatched = int(merged["ServiceCode"].isna().sum())
    merged["Region"] = clean_dimension(merged["Region"])
    merged["State"] = clean_dimension(merged["State"])
    merged["ServiceCenter"] = clean_dimension(merged["ServiceCenter"])
    
    report_df = merged.groupby(["Region", "State", "ServiceCenter"], as_index=False).agg(
        Unit=("PAYMENT STATUS", "count"), 
        GWP=("CUSTOMER PRICE", "sum")
    ).sort_values(["Region", "State", "ServiceCenter"], kind="stable")
    
    final_report = pd.DataFrame(build_final_rows(report_df))
    final_report.to_excel(FINAL_REPORT_FILE, index=False, sheet_name="Daily Report")
    style_service_workbook(FINAL_REPORT_FILE)
    
    grand_total = final_report[final_report["Region"] == "Grand Total"].iloc[0]
    
    return {
        "label": "Service",
        "summary": {
            "paid_rows": len(service),
            "total_gwp": int(grand_total["GWP"]),
            "regions": int(report_df["Region"].nunique()),
            "unmatched_rows": unmatched,
            "input_rows": input_rows,
            "service_centers": int(report_df["ServiceCenter"].nunique()),
            "total_units": int(grand_total["Unit"]),
            "generated_at": datetime.now().strftime("%d %b %Y, %I:%M %p")
        },
        "preview": final_report.head(40).to_dict(orient="records"),
        "columns": final_report.columns.tolist()
    }

def generate_channel_payload(axio_path: Path, retail_path: Path, master_path: Path) -> dict[str, object]:
    try:
        channel_report = generate_channel_report(
            axio_path=axio_path, 
            retail_path=retail_path, 
            master_path=master_path, 
            output_path=CHANNEL_REPORT_FILE
        )
        
        if channel_report.empty:
            raise ValueError("The generated channel report is empty. Check your input files.")
            
        grand_total_rows = channel_report[channel_report["State"] == "Grand Total"]
        if grand_total_rows.empty:
            # Fallback if Grand Total is named differently or missing
            grand_total = channel_report.iloc[-1]
        else:
            grand_total = grand_total_rows.iloc[0]
            
        # Safely count states and stores
        state_series = channel_report["State"].astype(str)
        state_count = state_series.str.endswith(" Total").sum()
        if "Grand Total" in state_series.values:
            state_count -= 1
            
        detail_rows = channel_report[
            ~state_series.str.contains("Total", na=False) & 
            ~channel_report["DistributorName"].astype(str).str.contains("Total", na=False)
        ]

        return {
            "label": "Retail + Axio",
            "summary": {
                "total_units": int(grand_total.get("Total Unit", 0)),
                "total_gwp": int(grand_total.get("Total GWP", 0)),
                "axio_units": int(grand_total.get("AXIO Unit", 0)),
                "retail_units": int(grand_total.get("Retail Unit", 0)),
                "states": int(max(0, state_count)),
                "stores": len(detail_rows),
                "generated_at": datetime.now().strftime("%d %b %Y, %I:%M %p")
            },
            "preview": channel_report.head(40).to_dict(orient="records"),
            "columns": channel_report.columns.tolist()
        }
    except Exception as e:
        raise ValueError(f"Channel Report Error: {str(e)}")

def file_status(path: Path | None, missing_name: str | None = None) -> dict[str, object]:
    if path is None:
        return {"exists": False, "name": missing_name or "Missing file"}
    if not path.exists(): return {"exists": False, "name": path.name}
    stat = path.stat()
    return {"exists": True, "name": path.name, "size": stat.st_size, "modified": datetime.fromtimestamp(stat.st_mtime).strftime("%d %b %Y, %I:%M %p")}

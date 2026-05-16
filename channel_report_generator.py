from __future__ import annotations

import os
import re
from pathlib import Path

import pandas as pd

from workbook_styles import style_channel_workbook

BASE_DIR = Path(__file__).resolve().parent
OUTPUT_DIR = Path("/tmp") if os.getenv("VERCEL") else BASE_DIR
CHANNEL_MASTER_PATTERNS = ("Master*.xlsb", "Master*.xlsx")
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


def first_existing(*names: str) -> Path:
    for name in names:
        path = BASE_DIR / name
        if path.exists():
            return path
    return BASE_DIR / names[0]


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


def resolve_master_file(base_dir: Path = BASE_DIR) -> Path:
    candidates: list[Path] = []
    for pattern in CHANNEL_MASTER_PATTERNS:
        candidates.extend(path for path in base_dir.glob(pattern) if path.is_file())

    if not candidates:
        return base_dir / "Master.xlsb"
    return sorted(candidates, key=channel_master_sort_key)[0]


AXIO_FILE = first_existing("axio.csv", "mi_smart_report (2).csv")
RETAIL_FILE = first_existing("retail copy.csv", "retail.csv", "mi_smart_report (1).csv")
MASTER_FILE = resolve_master_file()
OUTPUT_FILE = OUTPUT_DIR / "final_channel_report.xlsx"

VALUE_COLUMN = "Customer Price"
MASTER_LOOKUP_REQUIRED_COLUMNS = {"Retailer ID", "State", "Dist Name", "Outlet Name"}
PREFERRED_MASTER_SHEETS = (
    "Retail and Axio",
    "Retail & Axio",
    "Retail+Axio",
    "Retail_Axio",
    "Sheet1",
)


def clean_text(value: object) -> str:
    if pd.isna(value):
        return "Blank"
    cleaned = str(value).strip()
    if not cleaned or cleaned.lower() == "nan":
        return "Blank"
    return cleaned


def strip_column_names(frame: pd.DataFrame) -> pd.DataFrame:
    frame = frame.copy()
    frame.columns = frame.columns.str.strip()
    return frame


def normalise_sheet_name(value: str) -> str:
    return re.sub(r"[^a-z0-9]+", "", value.lower())


def ordered_with_blank_last(values: pd.Series) -> list[str]:
    unique_values = values.dropna().unique().tolist()
    ordered = [value for value in unique_values if value != "Blank"]
    if "Blank" in unique_values:
        ordered.append("Blank")
    return ordered


def resolve_master_lookup_sheet(master_path: Path) -> str | int:
    suffix = master_path.suffix.lower()
    engine = "pyxlsb" if suffix == ".xlsb" else None
    workbook = pd.ExcelFile(master_path, engine=engine)
    available_sheets = workbook.sheet_names
    normalised_map = {
        normalise_sheet_name(sheet_name): sheet_name for sheet_name in available_sheets
    }

    for preferred_name in PREFERRED_MASTER_SHEETS:
        matched_sheet = normalised_map.get(normalise_sheet_name(preferred_name))
        if matched_sheet:
            return matched_sheet

    for sheet_name in available_sheets:
        sample = strip_column_names(
            pd.read_excel(master_path, sheet_name=sheet_name, engine=engine, nrows=5)
        )
        if MASTER_LOOKUP_REQUIRED_COLUMNS.issubset(sample.columns):
            return sheet_name

    if available_sheets:
        return available_sheets[0]
    return 0


def read_master_lookup(master_path: Path = MASTER_FILE) -> pd.DataFrame:
    suffix = master_path.suffix.lower()
    engine = "pyxlsb" if suffix == ".xlsb" else None
    sheet_name = resolve_master_lookup_sheet(master_path)
    master = strip_column_names(
        pd.read_excel(master_path, sheet_name=sheet_name, engine=engine)
    )
    missing = sorted(MASTER_LOOKUP_REQUIRED_COLUMNS.difference(master.columns))
    if missing:
        raise ValueError(
            "Channel master sheet is missing required column(s): "
            + ", ".join(missing)
        )
    master["Retailer ID"] = pd.to_numeric(master["Retailer ID"], errors="coerce").astype("Int64")
    master["MasterState"] = master["State"].map(clean_text)
    master["MasterDistributorName"] = master["Dist Name"].map(clean_text)
    master["MasterRetailerName"] = master["Outlet Name"].map(clean_text)
    master = master.drop_duplicates(subset=["Retailer ID"])
    return master[
        ["Retailer ID", "MasterState", "MasterDistributorName", "MasterRetailerName"]
    ]


def filter_channel_rows(frame: pd.DataFrame, channel: str) -> pd.DataFrame:
    normalized_channel = clean_text(channel).upper()
    filtered = frame.copy()

    if normalized_channel == "RETAIL":
        if "Payment Status" in filtered.columns:
            filtered = filtered[
                filtered["Payment Status"].astype(str).str.upper().eq("TRUE")
            ].copy()
        if "Status" in filtered.columns:
            filtered = filtered[filtered["Status"].astype(str).ne("4")].copy()
    elif "Status" in filtered.columns:
        filtered = filtered[filtered["Status"].astype(str).ne("4")].copy()

    return filtered


def prepare_channel_frame(
    path: Path,
    channel: str,
    master_lookup: pd.DataFrame,
) -> pd.DataFrame:
    frame = strip_column_names(pd.read_csv(path))
    frame = filter_channel_rows(frame, channel)
    frame["Source"] = "Retail" if clean_text(channel).upper() == "RETAIL" else "Axio"
    frame["Retailer DMS id"] = pd.to_numeric(frame["Retailer DMS id"], errors="coerce").astype(
        "Int64"
    )
    frame[VALUE_COLUMN] = pd.to_numeric(frame[VALUE_COLUMN], errors="coerce")
    frame = frame.merge(
        master_lookup,
        left_on="Retailer DMS id",
        right_on="Retailer ID",
        how="left",
    )
    frame["Final_State"] = frame["MasterState"].fillna("Blank").map(clean_text)
    frame["Dist Name"] = frame["MasterDistributorName"].fillna("Blank").map(clean_text)
    frame["Outlet Name"] = frame["MasterRetailerName"].fillna("Blank").map(clean_text)
    return frame[["Final_State", "Dist Name", "Outlet Name", "Source", VALUE_COLUMN]]


def build_detail_report(combined: pd.DataFrame) -> pd.DataFrame:
    return (
        combined.groupby(
            ["Final_State", "Dist Name", "Outlet Name", "Source"],
            as_index=False,
        )
        .agg(
            Unit=(VALUE_COLUMN, "count"),
            GWP=(VALUE_COLUMN, "sum"),
        )
        .sort_values(
            ["Final_State", "Dist Name", "Outlet Name", "Source"],
            kind="stable",
        )
    )


def whole_number(value: object) -> int:
    if value in (None, "") or pd.isna(value):
        return 0
    return int(round(float(value)))


def build_final_rows(detail: pd.DataFrame) -> list[dict[str, object]]:
    rows: list[dict[str, object]] = []

    grand_axio_unit = 0
    grand_axio_gwp = 0
    grand_retail_unit = 0
    grand_retail_gwp = 0

    for state in ordered_with_blank_last(detail["Final_State"]):
        state_frame = detail[detail["Final_State"].eq(state)]

        state_axio_unit = 0
        state_axio_gwp = 0
        state_retail_unit = 0
        state_retail_gwp = 0
        first_state_row = True

        for distributor in ordered_with_blank_last(state_frame["Dist Name"]):
            distributor_frame = state_frame[state_frame["Dist Name"].eq(distributor)]

            dist_axio_unit = 0
            dist_axio_gwp = 0
            dist_retail_unit = 0
            dist_retail_gwp = 0
            first_distributor_row = True

            for retailer in ordered_with_blank_last(distributor_frame["Outlet Name"]):
                retailer_frame = distributor_frame[distributor_frame["Outlet Name"].eq(retailer)]
                axio_data = retailer_frame[retailer_frame["Source"].eq("Axio")]
                retail_data = retailer_frame[retailer_frame["Source"].eq("Retail")]

                axio_unit = whole_number(axio_data["Unit"].sum())
                axio_gwp = whole_number(axio_data["GWP"].sum())
                retail_unit = whole_number(retail_data["Unit"].sum())
                retail_gwp = whole_number(retail_data["GWP"].sum())
                total_unit = axio_unit + retail_unit
                total_gwp = axio_gwp + retail_gwp

                dist_axio_unit += axio_unit
                dist_axio_gwp += axio_gwp
                dist_retail_unit += retail_unit
                dist_retail_gwp += retail_gwp

                state_axio_unit += axio_unit
                state_axio_gwp += axio_gwp
                state_retail_unit += retail_unit
                state_retail_gwp += retail_gwp

                grand_axio_unit += axio_unit
                grand_axio_gwp += axio_gwp
                grand_retail_unit += retail_unit
                grand_retail_gwp += retail_gwp

                rows.append(
                    {
                        "State": state if first_state_row else "",
                        "DistributorName": distributor if first_distributor_row else "",
                        "RetailerName": retailer,
                        "AXIO Unit": axio_unit,
                        "AXIO GWP": axio_gwp,
                        "Retail Unit": retail_unit,
                        "Retail GWP": retail_gwp,
                        "Total Unit": total_unit,
                        "Total GWP": total_gwp,
                    }
                )

                first_state_row = False
                first_distributor_row = False

            rows.append(
                {
                    "State": "",
                    "DistributorName": f"{distributor} Total",
                    "RetailerName": "",
                    "AXIO Unit": dist_axio_unit,
                    "AXIO GWP": dist_axio_gwp,
                    "Retail Unit": dist_retail_unit,
                    "Retail GWP": dist_retail_gwp,
                    "Total Unit": dist_axio_unit + dist_retail_unit,
                    "Total GWP": dist_axio_gwp + dist_retail_gwp,
                }
            )

        rows.append(
            {
                "State": f"{state} Total",
                "DistributorName": "",
                "RetailerName": "",
                "AXIO Unit": state_axio_unit,
                "AXIO GWP": state_axio_gwp,
                "Retail Unit": state_retail_unit,
                "Retail GWP": state_retail_gwp,
                "Total Unit": state_axio_unit + state_retail_unit,
                "Total GWP": state_axio_gwp + state_retail_gwp,
            }
        )

    rows.append(
        {
            "State": "Grand Total",
            "DistributorName": "",
            "RetailerName": "",
            "AXIO Unit": grand_axio_unit,
            "AXIO GWP": grand_axio_gwp,
            "Retail Unit": grand_retail_unit,
            "Retail GWP": grand_retail_gwp,
            "Total Unit": grand_axio_unit + grand_retail_unit,
            "Total GWP": grand_axio_gwp + grand_retail_gwp,
        }
    )
    return rows


def generate_channel_report(
    axio_path: Path = AXIO_FILE,
    retail_path: Path = RETAIL_FILE,
    master_path: Path = MASTER_FILE,
    output_path: Path = OUTPUT_FILE,
) -> pd.DataFrame:
    master_lookup = read_master_lookup(master_path)
    axio = prepare_channel_frame(axio_path, "AXIO", master_lookup)
    retail = prepare_channel_frame(retail_path, "Retail", master_lookup)
    detail = build_detail_report(pd.concat([axio, retail], ignore_index=True))
    final_report = pd.DataFrame(build_final_rows(detail))
    numeric_columns = [
        "AXIO Unit",
        "AXIO GWP",
        "Retail Unit",
        "Retail GWP",
        "Total Unit",
        "Total GWP",
    ]
    final_report[numeric_columns] = final_report[numeric_columns].fillna(0).astype(int)
    final_report.to_excel(output_path, index=False)
    style_channel_workbook(output_path)
    return final_report


if __name__ == "__main__":
    report = generate_channel_report()
    grand_total = report[report["State"].eq("Grand Total")].iloc[0]
    print(
        "Generated final_channel_report.xlsx: "
        f"{int(grand_total['Total Unit'])} units, {int(grand_total['Total GWP'])} GWP"
    )

import argparse
import math
import os
import re
import unicodedata
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP
from datetime import date, datetime, time, timedelta
from pathlib import Path

import pandas as pd


DEFAULT_RAW_FILE = "C:\\Users\\cagatay\\Documents\\Ses Kay\u0131tlar\u0131\\data.xlsx"
DEFAULT_TC_CF_FILE = "C:\\Users\\cagatay\\Documents\\Ses Kay\u0131tlar\u0131\\cf.xlsx"
DEFAULT_LOGGER_CF_FILE = "C:\\Users\\cagatay\\Documents\\Ses Kay\u0131tlar\u0131\\cf2.xlsx"
INTERVAL_COUNT = 3
MAX_THERMOCOUPLE_COUNT = 40
ONE_DECIMAL = Decimal("0.1")


def choose_file_with_dialog(default: str | None = None) -> str | None:
    try:
        from tkinter import Tk, filedialog
    except ImportError:
        return None

    initial_dir = None
    initial_file = None
    if default:
        default_path = Path(default)
        if default_path.is_file():
            initial_dir = str(default_path.parent)
            initial_file = default_path.name
        elif default_path.parent.exists():
            initial_dir = str(default_path.parent)

    root = Tk()
    root.withdraw()
    root.attributes("-topmost", True)
    try:
        selected_path = filedialog.askopenfilename(
            title="Excel dosyasi secin",
            initialdir=initial_dir,
            initialfile=initial_file,
            filetypes=[
                ("Excel dosyalari", "*.xlsx *.xls"),
                ("Tum dosyalar", "*.*"),
            ],
        )
    finally:
        root.destroy()

    return selected_path or None


def sanitize_file_path(value: str) -> str:
    cleaned = str(value).strip().strip('"').strip("'")
    cleaned = "".join(
        char for char in cleaned if unicodedata.category(char) not in {"Cf", "Cc"}
    )
    return os.path.normpath(cleaned)


def ask_file_path(prompt: str, default: str | None = None) -> str:
    while True:
        if default:
            print(f"{prompt}:")
            print(f"  Enter = varsayilan dosya ({default})")
            print("  SEC   = dosya secme penceresi ac")
            path = input("  Yol: ")
        else:
            print(f"{prompt}:")
            print("  SEC = dosya secme penceresi ac")
            path = input("  Yol: ")

        normalized_input = sanitize_file_path(path)

        if not normalized_input and default:
            path = default
        elif normalized_input.upper() == "SEC":
            selected_path = choose_file_with_dialog(default)
            if selected_path:
                path = sanitize_file_path(selected_path)
                print(f"Secilen dosya: {path}")
            else:
                print("Dosya secilmedi. Tekrar deneyin.")
                continue
        else:
            path = normalized_input

        if os.path.isfile(path):
            return path
        print("Hata: Dosya bulunamadi. Tekrar deneyin.")


def ask_float(prompt: str, default: float | None = None) -> float:
    while True:
        shown_default = f" [{default}]" if default is not None else ""
        text = input(f"{prompt}{shown_default}: ").strip().replace(",", ".")
        if not text and default is not None:
            return default
        try:
            return float(text)
        except ValueError:
            print("Hata: Sayisal deger giriniz.")


def ask_required_time(prompt: str, default: str | None = None) -> time:
    while True:
        suffix = f" [{default}]" if default else ""
        text = input(f"{prompt}{suffix}: ").strip()
        if not text and default:
            text = default
        try:
            return parse_time_input(text)
        except ValueError:
            print("Hata: Saati HH:MM veya HH:MM:SS seklinde giriniz.")


def parse_time_input(text: str) -> time:
    cleaned = str(text).strip()
    for fmt in ("%H:%M", "%H:%M:%S"):
        try:
            return datetime.strptime(cleaned, fmt).time()
        except ValueError:
            continue
    raise ValueError("Saat formati HH:MM veya HH:MM:SS olmali.")


def ask_optional_time(prompt: str, default: str | None = None) -> time | None:
    while True:
        if default:
            suffix = f" [{default}]"
        else:
            suffix = " [tum veri icin bos birak]"

        text = input(f"{prompt}{suffix}: ").strip()
        if not text:
            if default:
                text = default
            else:
                return None

        try:
            return parse_time_input(text)
        except ValueError:
            print("Hata: Saati HH:MM veya HH:MM:SS seklinde giriniz.")


def ask_optional_int(prompt: str, default: int | None = None) -> int | None:
    while True:
        suffix = f" [{default}]" if default is not None else " [atlamak icin bos birak]"
        text = input(f"{prompt}{suffix}: ").strip()
        if not text:
            return default
        try:
            value = int(text)
        except ValueError:
            print("Hata: Tam sayi giriniz.")
            continue
        if value <= 0:
            print("Hata: 0'dan buyuk bir deger giriniz.")
            continue
        return value


def format_time_range(start_time: time, end_time: time) -> str:
    return f"{start_time.strftime('%H:%M:%S')} - {end_time.strftime('%H:%M:%S')}"


def parse_interval_arg(text: str, interval_index: int):
    parts = [part.strip() for part in str(text).split("|")]
    if len(parts) != 4:
        raise ValueError(
            f"Aralik {interval_index} formati 'baslangic|bitis|set noktasi|tolerans' olmali."
        )

    start_time = parse_time_input(parts[0])
    end_time = parse_time_input(parts[1])
    if end_time < start_time:
        raise ValueError(f"Aralik {interval_index} bitis saati baslangictan kucuk olamaz.")

    try:
        setpoint = float(parts[2].replace(",", "."))
        tolerance = float(parts[3].replace(",", "."))
    except ValueError as exc:
        raise ValueError(f"Aralik {interval_index} set noktasi ve tolerans sayisal olmali.") from exc

    if tolerance < 0:
        raise ValueError(f"Aralik {interval_index} tolerans negatif olamaz.")

    return {
        "index": interval_index,
        "label": f"ARALIK {interval_index}",
        "start_time": start_time,
        "end_time": end_time,
        "setpoint": setpoint,
        "tolerance": tolerance,
    }


def ask_interval_configs(interval_count: int = INTERVAL_COUNT):
    interval_configs = []

    for interval_index in range(1, interval_count + 1):
        while True:
            print("-" * 100)
            print(f"{interval_index}. SET NOKTASI ARALIGI")
            print("-" * 100)
            start_time = ask_required_time(f"Aralik {interval_index} baslangic saati")
            end_time = ask_required_time(f"Aralik {interval_index} bitis saati")
            if end_time < start_time:
                print("Hata: Bitis saati baslangic saatinden kucuk olamaz. Bu araligi tekrar girin.")
                continue

            setpoint = ask_float(f"Aralik {interval_index} set noktasi")
            tolerance = ask_float(f"Aralik {interval_index} toleransi")
            if tolerance < 0:
                print("Hata: Tolerans negatif olamaz. Bu araligi tekrar girin.")
                continue

            interval_configs.append(
                {
                    "index": interval_index,
                    "label": f"ARALIK {interval_index}",
                    "start_time": start_time,
                    "end_time": end_time,
                    "setpoint": setpoint,
                    "tolerance": tolerance,
                }
            )
            break

    return interval_configs


def normalize_text(value) -> str:
    return (
        str(value)
        .strip()
        .replace("\u0130", "I")
        .replace("\u0131", "i")
        .replace(" ", "")
        .replace("_", "")
        .replace("-", "")
        .upper()
    )


def strip_pandas_duplicate_suffix(value) -> str:
    return re.sub(r"\.\d+$", "", str(value).strip())


def excel_column_letter(column_index: int) -> str:
    letters = []
    while column_index > 0:
        column_index, remainder = divmod(column_index - 1, 26)
        letters.append(chr(65 + remainder))
    return "".join(reversed(letters))


def canonical_tc_name(value) -> str:
    text = normalize_text(value)
    match = re.fullmatch(r"(TC)?(\d+)", text)
    if match:
        return f"TC{int(match.group(2))}"
    return text


def get_tc_number_from_name(value) -> int | None:
    tc_name = canonical_tc_name(value)
    match = re.fullmatch(r"TC(\d+)", tc_name)
    if not match:
        return None
    return int(match.group(1))


def tc_sort_key(value):
    tc_number = get_tc_number_from_name(value)
    if tc_number is not None:
        return (0, tc_number, "")
    return (1, 0, canonical_tc_name(value))


def is_blank(value) -> bool:
    if value is None:
        return True
    if isinstance(value, float) and math.isnan(value):
        return True
    return str(value).strip() == ""


def is_valid_number(value) -> bool:
    if is_blank(value):
        return False
    try:
        Decimal(str(value).replace(",", "."))
        return True
    except (InvalidOperation, ValueError):
        return False


def round_to_one_decimal(value) -> float:
    decimal_value = Decimal(str(value).replace(",", "."))
    return float(decimal_value.quantize(ONE_DECIMAL, rounding=ROUND_HALF_UP))


def to_float(value) -> float:
    return round_to_one_decimal(value)


def format_number(value: float) -> str:
    return f"{float(value):.1f}"


def format_signed_number(value: float) -> str:
    return f"{float(value):+.1f}"


def extract_setpoint_value(value) -> float | None:
    if is_blank(value):
        return None

    if isinstance(value, (int, float)):
        try:
            return round_to_one_decimal(value)
        except (InvalidOperation, ValueError):
            return None

    text = strip_pandas_duplicate_suffix(value)
    if not text:
        return None

    normalized = normalize_text(text)
    if normalized in {
        "CF",
        "CORRECTIONFACTOR",
        "CORRECTION",
        "CORRECTIONFACTORS",
        "OFFSET",
    }:
        return None

    match = re.search(r"[-+]?\d+(?:[.,]\d+)?", text)
    if not match:
        return None

    try:
        return round_to_one_decimal(match.group(0))
    except (InvalidOperation, ValueError):
        return None


def is_cf_like_column(value) -> bool:
    return normalize_text(strip_pandas_duplicate_suffix(value)) in {
        "CF",
        "CORRECTIONFACTOR",
        "CORRECTION",
        "CORRECTIONFACTORS",
        "OFFSET",
    }


def find_column(df: pd.DataFrame, alternatives):
    col_map = {normalize_text(col): col for col in df.columns if not is_blank(col)}
    for alt in alternatives:
        norm = normalize_text(alt)
        if norm in col_map:
            return col_map[norm]
    return None


def format_axis_value(value) -> str:
    if isinstance(value, pd.Timestamp):
        value = value.to_pydatetime()

    if isinstance(value, datetime):
        return value.strftime("%H:%M:%S") if value.date() == date(1900, 1, 1) else value.strftime("%Y-%m-%d %H:%M:%S")
    if isinstance(value, time):
        return value.strftime("%H:%M:%S")
    if isinstance(value, date):
        return value.strftime("%Y-%m-%d")

    return str(value)


def format_time_for_display(value: time | None) -> str:
    if value is None:
        return "Tum Veri"
    return value.strftime("%H:%M:%S")


def format_evaluation_window(value: time | None) -> str:
    if value is None:
        return "Tum veri"
    return f"{value.strftime('%H:%M:%S')} sonrasi"


def format_time_phrase(value: time | None) -> str:
    if value is None:
        return "Tum veride"
    if value.second == 0:
        return f"{value.strftime('%H:%M')} sonrasinda"
    return f"{value.strftime('%H:%M:%S')} sonrasinda"


def format_minutes_for_display(value: int | None) -> str:
    if value is None:
        return "Yapilmadi"
    return f"{value} dk"


def excel_serial_to_datetime(numeric: float) -> datetime | None:
    if math.isnan(numeric) or numeric < 0:
        return None
    excel_origin = datetime(1899, 12, 30)
    return excel_origin + timedelta(days=numeric)


def extract_time_of_day(value) -> time | None:
    if is_blank(value):
        return None

    if isinstance(value, pd.Timestamp):
        value = value.to_pydatetime()

    if isinstance(value, datetime):
        return value.time().replace(microsecond=0)
    if isinstance(value, time):
        return value.replace(microsecond=0)

    if isinstance(value, (int, float)):
        numeric = float(value)
        excel_dt = excel_serial_to_datetime(numeric)
        if excel_dt is not None:
            total_seconds = int(round((numeric % 1) * 24 * 60 * 60))
            total_seconds %= 24 * 60 * 60
            hours, remainder = divmod(total_seconds, 3600)
            minutes, seconds = divmod(remainder, 60)
            return time(hours, minutes, seconds)

    text = str(value).strip()
    if not text:
        return None

    for fmt in (
        "%H:%M",
        "%H:%M:%S",
        "%H.%M",
        "%H.%M.%S",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d %H:%M:%S",
        "%d.%m.%Y %H:%M",
        "%d.%m.%Y %H:%M:%S",
        "%d/%m/%Y %H:%M",
        "%d/%m/%Y %H:%M:%S",
    ):
        try:
            return datetime.strptime(text, fmt).time()
        except ValueError:
            continue

    parsed = pd.to_datetime(text, errors="coerce", dayfirst=True)
    if not pd.isna(parsed):
        return parsed.to_pydatetime().time().replace(microsecond=0)

    return None


def extract_comparable_datetime(value) -> datetime | None:
    if is_blank(value):
        return None

    if isinstance(value, pd.Timestamp):
        value = value.to_pydatetime()

    if isinstance(value, datetime):
        return value.replace(microsecond=0)
    if isinstance(value, time):
        return datetime.combine(date(1900, 1, 1), value.replace(microsecond=0))
    if isinstance(value, date):
        return datetime.combine(value, time(0, 0, 0))

    if isinstance(value, (int, float)):
        numeric = float(value)
        excel_dt = excel_serial_to_datetime(numeric)
        if excel_dt is not None:
            return excel_dt.replace(microsecond=0)

    text = str(value).strip()
    if not text:
        return None

    for fmt in (
        "%H:%M",
        "%H:%M:%S",
        "%H.%M",
        "%H.%M.%S",
        "%Y-%m-%d %H:%M",
        "%Y-%m-%d %H:%M:%S",
        "%d.%m.%Y %H:%M",
        "%d.%m.%Y %H:%M:%S",
        "%d/%m/%Y %H:%M",
        "%d/%m/%Y %H:%M:%S",
    ):
        try:
            parsed = datetime.strptime(text, fmt)
            if fmt in ("%H:%M", "%H:%M:%S"):
                return datetime.combine(date(1900, 1, 1), parsed.time())
            return parsed
        except ValueError:
            continue

    parsed = pd.to_datetime(text, errors="coerce", dayfirst=True)
    if not pd.isna(parsed):
        return parsed.to_pydatetime().replace(microsecond=0)

    return None


def build_evaluation_row_mask(corrected_df: pd.DataFrame, evaluation_start_time: time | None):
    if evaluation_start_time is None:
        return [True] * len(corrected_df), []

    time_col = corrected_df.columns[0]
    selected_rows = []
    invalid_rows = []

    for row_index, value in enumerate(corrected_df[time_col].tolist(), start=2):
        row_time = extract_time_of_day(value)
        if row_time is None:
            invalid_rows.append(str(row_index))
            selected_rows.append(False)
            continue
        selected_rows.append(row_time >= evaluation_start_time)

    return selected_rows, invalid_rows


def build_interval_row_mask(
    corrected_df: pd.DataFrame, start_time: time, end_time: time
):
    time_col = corrected_df.columns[0]
    selected_rows = []
    invalid_rows = []

    for row_index, value in enumerate(corrected_df[time_col].tolist(), start=2):
        row_time = extract_time_of_day(value)
        if row_time is None:
            invalid_rows.append(str(row_index))
            selected_rows.append(False)
            continue
        selected_rows.append(start_time <= row_time <= end_time)

    return selected_rows, invalid_rows


def make_output_paths(base_file: str, output_dir: str | None):
    target_dir = Path(output_dir) if output_dir else Path(base_file).resolve().parent
    target_dir.mkdir(parents=True, exist_ok=True)

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    excel_path = target_dir / f"tus_corrected_output_{timestamp}.xlsx"
    report_path = target_dir / f"tus_report_{timestamp}.txt"
    full_chart_path = target_dir / f"tus_all_intervals_chart_{timestamp}.png"
    interval_chart_paths = [
        target_dir / f"tus_interval_{index}_chart_{timestamp}.png"
        for index in range(1, INTERVAL_COUNT + 1)
    ]
    return excel_path, report_path, full_chart_path, interval_chart_paths


def read_excel_safely(file_path: str, label: str) -> pd.DataFrame:
    try:
        df = pd.read_excel(file_path)
    except FileNotFoundError as exc:
        raise FileNotFoundError(f"{label} dosyasi bulunamadi: {file_path}") from exc
    except Exception as exc:
        raise ValueError(f"{label} dosyasi okunamadi: {exc}") from exc

    if df.empty:
        raise ValueError(f"{label} dosyasi bos.")

    df = df.dropna(axis=1, how="all").dropna(axis=0, how="all")
    if df.empty:
        raise ValueError(f"{label} dosyasinda kullanilabilir veri yok.")

    return df


def load_raw_data(raw_file: str):
    df = read_excel_safely(raw_file, "Ham veri")

    if df.shape[1] < 2:
        raise ValueError("Ham veri dosyasinda en az 2 sutun olmali.")

    time_col = df.columns[0]
    raw_columns = list(df.columns[1:])
    time_values = df[time_col].tolist()
    raw_tc_data = {}

    for col in raw_columns:
        if df[col].isna().all():
            continue

        tc_name = canonical_tc_name(col)
        if tc_name in raw_tc_data:
            raise ValueError(f"Ham veri dosyasinda tekrar eden TC kolonu var: {tc_name}")

        values = []
        for row_index, cell_value in enumerate(df[col], start=2):
            if is_blank(cell_value):
                raise ValueError(
                    f"Ham veri dosyasinda '{col}' sutununda bos hucre var. Satir: {row_index}"
                )
            if is_valid_number(cell_value):
                values.append(to_float(cell_value))
            else:
                raise ValueError(
                    f"Ham veri dosyasinda '{col}' sutununda gecersiz veri var. Satir: {row_index}"
                )

        if not values:
            raise ValueError(f"Ham veri dosyasinda '{col}' sutununda hic sayisal veri yok.")

        raw_tc_data[tc_name] = values

    if not raw_tc_data:
        raise ValueError("Ham veri dosyasinda islenecek TC kolonu bulunamadi.")
    if len(raw_tc_data) > MAX_THERMOCOUPLE_COUNT:
        raise ValueError(
            f"Ham veri dosyasinda {len(raw_tc_data)} adet thermocouple bulundu. "
            f"Program en fazla {MAX_THERMOCOUPLE_COUNT} thermocouple destekler."
        )

    return time_col, time_values, raw_tc_data


def load_cf_data(cf_file: str, label: str):
    df = read_excel_safely(cf_file, label)

    tc_col = find_column(df, ["TC", "Thermocouple", "Point", "Nokta", "Sensor"])

    if tc_col is None:
        raise ValueError(f"{label} dosyasinda TC kolonu bulunamadi.")
    column_entries = []
    seen_setpoints = {}

    for column_index, col in enumerate(df.columns, start=1):
        if col == tc_col or df[col].isna().all():
            continue

        column_label = excel_column_letter(column_index)
        base_name = str(strip_pandas_duplicate_suffix(col))
        column_entry = {
            "column": col,
            "display_name": f"{base_name} ({column_label})",
            "map": {},
        }

        if is_cf_like_column(col):
            column_entry["kind"] = "generic"
            column_entries.append(column_entry)
            continue

        setpoint_value = extract_setpoint_value(col)
        if setpoint_value is None:
            continue
        if setpoint_value in seen_setpoints:
            raise ValueError(
                f"{label} dosyasinda ayni set noktasi icin birden fazla CF kolonu var: "
                f"{format_number(setpoint_value)}"
            )

        seen_setpoints[setpoint_value] = col
        column_entry["kind"] = "setpoint"
        column_entry["setpoint"] = setpoint_value
        column_entries.append(column_entry)

    if not column_entries:
        raise ValueError(
            f"{label} dosyasinda CF kolonu bulunamadi. "
            "Tek bir 'CF' kolonu ya da set noktasi adli CF kolonlari bekleniyor."
        )

    for row_index, row in df.iterrows():
        tc_value = row[tc_col]

        if is_blank(tc_value):
            raise ValueError(f"{label} dosyasinda {row_index + 2}. satirda TC bos.")

        tc_name = canonical_tc_name(tc_value)

        for entry in column_entries:
            cf_value = row[entry["column"]]
            if not is_valid_number(cf_value):
                raise ValueError(
                    f"{label} dosyasinda {row_index + 2}. satirda "
                    f"'{entry['column']}' kolonu icin CF gecersiz."
                )
            if tc_name in entry["map"]:
                raise ValueError(
                    f"{label} dosyasinda tekrar eden TC kaydi var: {tc_name}"
                )
            entry["map"][tc_name] = to_float(cf_value)

    return {
        "label": label,
        "ordered_columns": column_entries,
        "setpoints": {
            entry["setpoint"]: entry
            for entry in column_entries
            if entry["kind"] == "setpoint"
        },
    }


def resolve_cf_columns_for_intervals(cf_data, interval_configs):
    unique_setpoints = []
    for interval_config in interval_configs:
        setpoint_key = round_to_one_decimal(interval_config["setpoint"])
        if setpoint_key not in unique_setpoints:
            unique_setpoints.append(setpoint_key)

    resolved = {}
    used_columns = set()

    for setpoint_key in unique_setpoints:
        if setpoint_key in cf_data["setpoints"]:
            entry = cf_data["setpoints"][setpoint_key]
            resolved[setpoint_key] = entry
            used_columns.add(entry["column"])

    unresolved_setpoints = [value for value in unique_setpoints if value not in resolved]
    remaining_columns = [
        entry for entry in cf_data["ordered_columns"] if entry["column"] not in used_columns
    ]

    if unresolved_setpoints:
        if len(cf_data["ordered_columns"]) == 1:
            entry = cf_data["ordered_columns"][0]
            for setpoint_key in unresolved_setpoints:
                resolved[setpoint_key] = entry
        elif len(remaining_columns) == len(unresolved_setpoints):
            for setpoint_key, entry in zip(unresolved_setpoints, remaining_columns):
                resolved[setpoint_key] = entry
        else:
            available_names = ", ".join(
                entry["display_name"] for entry in cf_data["ordered_columns"]
            )
            raise ValueError(
                f"{cf_data['label']} dosyasinda set noktalarina uygun CF kolonlari secilemedi. "
                f"Istenen set noktalar: {', '.join(format_number(value) for value in unique_setpoints)} | "
                f"Mevcut kolonlar: {available_names}"
            )

    return resolved


def validate_tc_coverage(raw_tc_data, tc_cf_map, logger_cf_map):
    raw_tcs = set(raw_tc_data)
    tc_cf_tcs = set(tc_cf_map)
    logger_cf_tcs = set(logger_cf_map)

    missing_tc_cf = sorted(raw_tcs - tc_cf_tcs, key=tc_sort_key)
    missing_logger_cf = sorted(raw_tcs - logger_cf_tcs, key=tc_sort_key)
    extra_tc_cf = sorted(tc_cf_tcs - raw_tcs, key=tc_sort_key)
    extra_logger_cf = sorted(logger_cf_tcs - raw_tcs, key=tc_sort_key)

    errors = []
    if missing_tc_cf:
        errors.append(f"Thermocouple CF dosyasinda eksik TC: {', '.join(missing_tc_cf)}")
    if missing_logger_cf:
        errors.append(f"Datalogger CF dosyasinda eksik TC: {', '.join(missing_logger_cf)}")
    if extra_tc_cf:
        errors.append(f"Thermocouple CF dosyasinda fazla TC: {', '.join(extra_tc_cf)}")
    if extra_logger_cf:
        errors.append(f"Datalogger CF dosyasinda fazla TC: {', '.join(extra_logger_cf)}")

    if errors:
        raise ValueError(" | ".join(errors))


def build_corrected_data(time_col, time_values, raw_tc_data, tc_cf_map, logger_cf_map):
    corrected_output = {str(time_col): time_values}

    for tc_name in sorted(raw_tc_data, key=tc_sort_key):
        raw_values = raw_tc_data[tc_name]
        tc_cf = tc_cf_map[tc_name]
        logger_cf = logger_cf_map[tc_name]

        corrected_values = [round_to_one_decimal(v + tc_cf + logger_cf) for v in raw_values]
        corrected_output[tc_name] = corrected_values

    corrected_df = pd.DataFrame(corrected_output)
    return corrected_df


def build_raw_data_frame(time_col, time_values, raw_tc_data):
    raw_output = {str(time_col): time_values}
    for tc_name in sorted(raw_tc_data, key=tc_sort_key):
        raw_output[tc_name] = raw_tc_data[tc_name]
    return pd.DataFrame(raw_output)


def prepare_interval_corrections(raw_tc_data, tc_cf_data, logger_cf_data, interval_configs):
    interval_corrections = []
    tc_cf_resolved = resolve_cf_columns_for_intervals(tc_cf_data, interval_configs)
    logger_cf_resolved = resolve_cf_columns_for_intervals(logger_cf_data, interval_configs)

    for interval_config in interval_configs:
        setpoint_key = round_to_one_decimal(interval_config["setpoint"])
        tc_cf_entry = tc_cf_resolved[setpoint_key]
        logger_cf_entry = logger_cf_resolved[setpoint_key]
        tc_cf_map = tc_cf_entry["map"]
        logger_cf_map = logger_cf_entry["map"]
        tc_cf_column = tc_cf_entry["display_name"]
        logger_cf_column = logger_cf_entry["display_name"]
        validate_tc_coverage(raw_tc_data, tc_cf_map, logger_cf_map)
        interval_corrections.append(
            {
                "config": interval_config,
                "tc_cf_map": tc_cf_map,
                "logger_cf_map": logger_cf_map,
                "tc_cf_column": tc_cf_column,
                "logger_cf_column": logger_cf_column,
            }
        )

    return interval_corrections


def build_combined_corrected_data(time_col, time_values, raw_tc_data, interval_corrections):
    tc_names = sorted(raw_tc_data, key=tc_sort_key)
    corrected_output = {str(time_col): time_values}
    row_assignments = [None] * len(time_values)

    for tc_name in tc_names:
        corrected_output[tc_name] = [math.nan] * len(time_values)

    time_df = pd.DataFrame({str(time_col): time_values})

    for interval_correction in interval_corrections:
        config = interval_correction["config"]
        selected_rows, _ = build_interval_row_mask(
            time_df, config["start_time"], config["end_time"]
        )

        for row_index, is_selected in enumerate(selected_rows):
            if not is_selected:
                continue
            if row_assignments[row_index] is not None:
                raise ValueError(
                    f"{config['label']} araligindeki bir satir baska bir aralikla cakismaktadir. "
                    "Aralik saatlerini tekrar kontrol edin."
                )

            row_assignments[row_index] = config["label"]
            for tc_name in tc_names:
                raw_value = raw_tc_data[tc_name][row_index]
                corrected_output[tc_name][row_index] = round_to_one_decimal(
                    raw_value
                    + interval_correction["tc_cf_map"][tc_name]
                    + interval_correction["logger_cf_map"][tc_name]
                )

    return pd.DataFrame(corrected_output)


def build_chart_display_data(raw_df: pd.DataFrame, corrected_df: pd.DataFrame):
    display_df = raw_df.copy()

    for tc_name in corrected_df.columns[1:]:
        corrected_values = corrected_df[tc_name]
        valid_mask = corrected_values.notna()
        display_df.loc[valid_mask, tc_name] = corrected_values.loc[valid_mask]

    return display_df


def filter_corrected_data_by_time(corrected_df: pd.DataFrame, evaluation_start_time: time | None):
    selected_rows, invalid_rows = build_evaluation_row_mask(corrected_df, evaluation_start_time)

    filtered_df = corrected_df.loc[selected_rows].reset_index(drop=True)
    if filtered_df.empty:
        if invalid_rows:
            rows_text = ", ".join(invalid_rows[:10])
            if len(invalid_rows) > 10:
                rows_text += ", ..."
            raise ValueError(
                f"{format_time_for_display(evaluation_start_time)} saatinden sonra uygun veri bulunamadi. "
                f"Saati okunamayan satirlar: {rows_text}"
            )
        raise ValueError(
            f"{format_time_for_display(evaluation_start_time)} saatinden sonra veri bulunamadi."
        )

    return filtered_df, invalid_rows


def filter_corrected_data_by_interval(
    corrected_df: pd.DataFrame, start_time: time, end_time: time
):
    selected_rows, invalid_rows = build_interval_row_mask(corrected_df, start_time, end_time)
    selected_row_indices = [index for index, keep in enumerate(selected_rows) if keep]
    filtered_df = corrected_df.loc[selected_rows].reset_index(drop=True)

    if filtered_df.empty:
        if invalid_rows:
            rows_text = ", ".join(invalid_rows[:10])
            if len(invalid_rows) > 10:
                rows_text += ", ..."
            raise ValueError(
                f"{format_time_range(start_time, end_time)} araliginda uygun veri bulunamadi. "
                f"Saati okunamayan satirlar: {rows_text}"
            )
        raise ValueError(
            f"{format_time_range(start_time, end_time)} araliginda veri bulunamadi."
        )

    return filtered_df, invalid_rows, selected_row_indices


def summarize_corrected_data(corrected_df, tc_cf_map, logger_cf_map):
    tc_columns = list(corrected_df.columns[1:])
    if not tc_columns:
        raise ValueError("Duzeltilmis veri icinde TC kolonu bulunamadi.")

    summary_rows = []
    all_corrected_values = []

    for tc_name in tc_columns:
        if tc_name not in tc_cf_map:
            raise ValueError(f"Thermocouple CF dosyasinda eksik TC var: {tc_name}")
        if tc_name not in logger_cf_map:
            raise ValueError(f"Datalogger CF dosyasinda eksik TC var: {tc_name}")

        corrected_values = [float(v) for v in corrected_df[tc_name].tolist()]
        tc_cf = tc_cf_map[tc_name]
        logger_cf = logger_cf_map[tc_name]
        total_cf = round_to_one_decimal(tc_cf + logger_cf)
        raw_values = [round_to_one_decimal(v - total_cf) for v in corrected_values]

        point_min = min(corrected_values)
        point_max = max(corrected_values)
        point_avg = round_to_one_decimal(sum(corrected_values) / len(corrected_values))

        summary_rows.append(
            {
                "TC": tc_name,
                "RAW_MIN": min(raw_values),
                "RAW_MAX": max(raw_values),
                "TC_CF": tc_cf,
                "DATALOGGER_CF": logger_cf,
                "TOTAL_CF": total_cf,
                "MIN_CORRECTED": point_min,
                "MAX_CORRECTED": point_max,
                "AVG_CORRECTED": point_avg,
            }
        )

        all_corrected_values.extend(corrected_values)

    summary_df = (
        pd.DataFrame(summary_rows)
        .assign(TC_SORT_KEY=lambda frame: frame["TC"].map(tc_sort_key))
        .sort_values("TC_SORT_KEY")
        .drop(columns=["TC_SORT_KEY"])
        .reset_index(drop=True)
    )
    return summary_df, all_corrected_values


def find_extreme_points_in_window(corrected_df: pd.DataFrame, evaluation_start_time: time | None):
    if corrected_df.empty:
        raise ValueError("En sicak ve en soguk noktalar icin veri bulunamadi.")

    time_col = corrected_df.columns[0]
    tc_columns = list(corrected_df.columns[1:])
    if not tc_columns:
        raise ValueError("En sicak ve en soguk noktalar icin TC kolonu bulunamadi.")

    row_mask, _ = build_evaluation_row_mask(corrected_df, evaluation_start_time)
    selected_row_indices = [index for index, keep in enumerate(row_mask) if keep]
    if not selected_row_indices:
        raise ValueError("Secilen degerlendirme araliginda veri bulunamadi.")

    hottest_info = None
    coldest_info = None

    for window_index, row_index in enumerate(selected_row_indices):
        row = corrected_df.iloc[row_index]
        time_label = format_axis_value(row[time_col])

        for tc_name in tc_columns:
            value = float(row[tc_name])

            hottest_candidate = {
                "hottest_tc": tc_name,
                "hottest_tc_number": get_tc_number_from_name(tc_name),
                "hottest_value": value,
                "hottest_time_label": time_label,
                "hottest_row_index": row_index,
                "hottest_window_index": window_index,
            }
            if hottest_info is None or value > hottest_info["hottest_value"]:
                hottest_info = hottest_candidate

            coldest_candidate = {
                "coldest_tc": tc_name,
                "coldest_tc_number": get_tc_number_from_name(tc_name),
                "coldest_value": value,
                "coldest_time_label": time_label,
                "coldest_row_index": row_index,
                "coldest_window_index": window_index,
            }
            if coldest_info is None or value < coldest_info["coldest_value"]:
                coldest_info = coldest_candidate

    return {**hottest_info, **coldest_info}


def find_extreme_points_in_rows(corrected_df: pd.DataFrame, selected_row_indices):
    if not selected_row_indices:
        raise ValueError("En sicak ve en soguk noktalar icin secili satir bulunamadi.")

    time_col = corrected_df.columns[0]
    tc_columns = list(corrected_df.columns[1:])
    if not tc_columns:
        raise ValueError("En sicak ve en soguk noktalar icin TC kolonu bulunamadi.")

    hottest_info = None
    coldest_info = None

    for local_index, row_index in enumerate(selected_row_indices):
        row = corrected_df.iloc[row_index]
        time_label = format_axis_value(row[time_col])

        for tc_name in tc_columns:
            value = float(row[tc_name])

            hottest_candidate = {
                "hottest_tc": tc_name,
                "hottest_tc_number": get_tc_number_from_name(tc_name),
                "hottest_value": value,
                "hottest_time_label": time_label,
                "hottest_row_index": row_index,
                "hottest_window_index": local_index,
            }
            if hottest_info is None or value > hottest_info["hottest_value"]:
                hottest_info = hottest_candidate

            coldest_candidate = {
                "coldest_tc": tc_name,
                "coldest_tc_number": get_tc_number_from_name(tc_name),
                "coldest_value": value,
                "coldest_time_label": time_label,
                "coldest_row_index": row_index,
                "coldest_window_index": local_index,
            }
            if coldest_info is None or value < coldest_info["coldest_value"]:
                coldest_info = coldest_candidate

    return {**hottest_info, **coldest_info}


def evaluate(summary_df, all_corrected_values, setpoint, tolerance):
    if tolerance < 0:
        raise ValueError("Tolerans negatif olamaz.")
    if not all_corrected_values:
        raise ValueError("Degerlendirme icin duzeltilmis veri bulunamadi.")

    allowed_min = setpoint - tolerance
    allowed_max = setpoint + tolerance

    overall_min = round_to_one_decimal(min(all_corrected_values))
    overall_max = round_to_one_decimal(max(all_corrected_values))
    spread = round_to_one_decimal(overall_max - overall_min)

    point_results = []
    for _, row in summary_df.iterrows():
        point_pass = (
            row["MIN_CORRECTED"] >= allowed_min and row["MAX_CORRECTED"] <= allowed_max
        )
        point_results.append("PASS" if point_pass else "FAIL")

    summary_df["POINT_RESULT"] = point_results

    overall_result = "PASS"
    if any(v < allowed_min or v > allowed_max for v in all_corrected_values):
        overall_result = "FAIL"

    hottest_index = summary_df["MAX_CORRECTED"].idxmax()
    coldest_index = summary_df["MIN_CORRECTED"].idxmin()

    result_info = {
        "allowed_min": allowed_min,
        "allowed_max": allowed_max,
        "overall_min": overall_min,
        "overall_max": overall_max,
        "spread": spread,
        "overall_result": overall_result,
        "overall_result_reason": "",
        "hottest_tc": summary_df.loc[hottest_index, "TC"],
        "hottest_tc_number": get_tc_number_from_name(summary_df.loc[hottest_index, "TC"]),
        "hottest_value": round_to_one_decimal(summary_df.loc[hottest_index, "MAX_CORRECTED"]),
        "coldest_tc": summary_df.loc[coldest_index, "TC"],
        "coldest_tc_number": get_tc_number_from_name(summary_df.loc[coldest_index, "TC"]),
        "coldest_value": round_to_one_decimal(summary_df.loc[coldest_index, "MIN_CORRECTED"]),
    }

    return summary_df, result_info


def analyze_stabilization(evaluation_df, setpoint, tolerance, stabilization_window_minutes):
    if stabilization_window_minutes is None:
        return None

    time_col = evaluation_df.columns[0]
    tc_columns = list(evaluation_df.columns[1:])
    if evaluation_df.empty:
        return {
            "window_minutes": stabilization_window_minutes,
            "status": "BULUNAMADI",
            "message": "Degerlendirilecek veri bulunamadi.",
        }

    allowed_min = setpoint - tolerance
    allowed_max = setpoint + tolerance
    row_times = []
    row_pass = []
    invalid_rows = []

    for row_index, (_, row) in enumerate(evaluation_df.iterrows(), start=2):
        comparable_dt = extract_comparable_datetime(row[time_col])
        if comparable_dt is None:
            invalid_rows.append(str(row_index))
            row_times.append(None)
            row_pass.append(False)
            continue

        row_values = [float(row[tc_name]) for tc_name in tc_columns]
        row_times.append(comparable_dt)
        row_pass.append(all(allowed_min <= value <= allowed_max for value in row_values))

    if invalid_rows:
        valid_indices = [i for i, dt in enumerate(row_times) if dt is not None]
        if not valid_indices:
            rows_text = ", ".join(invalid_rows[:10])
            if len(invalid_rows) > 10:
                rows_text += ", ..."
            raise ValueError(
                f"Stabilizasyon kontrolu icin saat kolonu okunamadi. Problemli satirlar: {rows_text}"
            )

    required_delta = timedelta(minutes=stabilization_window_minutes)

    for start_index, start_dt in enumerate(row_times):
        if start_dt is None or not row_pass[start_index]:
            continue

        end_index = None
        all_rows_in_window_pass = True

        for current_index in range(start_index, len(row_times)):
            current_dt = row_times[current_index]
            if current_dt is None:
                all_rows_in_window_pass = False
                break

            if not row_pass[current_index]:
                all_rows_in_window_pass = False
                break

            if current_dt - start_dt >= required_delta:
                end_index = current_index
                break

        if all_rows_in_window_pass and end_index is not None:
            return {
                "window_minutes": stabilization_window_minutes,
                "status": "BULUNDU",
                "start_label": format_axis_value(evaluation_df.iloc[start_index, 0]),
                "end_label": format_axis_value(evaluation_df.iloc[end_index, 0]),
                "row_count": end_index - start_index + 1,
            }

    return {
        "window_minutes": stabilization_window_minutes,
        "status": "BULUNAMADI",
        "message": "Belirtilen dakika boyunca tum noktalar tolerans icinde kalmadi.",
    }


def analyze_full_data_overshoot(corrected_df: pd.DataFrame, setpoint: float, tolerance: float):
    if corrected_df.empty:
        raise ValueError("Tum veri overshoot kontrolu icin veri bulunamadi.")

    time_col = corrected_df.columns[0]
    tc_columns = list(corrected_df.columns[1:])
    allowed_max = setpoint + tolerance

    overshoot_points = []

    for row_index, (_, row) in enumerate(corrected_df.iterrows(), start=2):
        time_label = format_axis_value(row[time_col])

        for tc_name in tc_columns:
            value = float(row[tc_name])
            if value > allowed_max:
                overshoot_points.append(
                    {
                        "row_number": row_index,
                        "time_label": time_label,
                        "tc_name": tc_name,
                        "tc_number": get_tc_number_from_name(tc_name),
                        "value": value,
                    }
                )

    result = {
        "overshoot_result": "VAR" if overshoot_points else "YOK",
        "overshoot_point_count": len(overshoot_points),
    }

    if overshoot_points:
        first_point = overshoot_points[0]
        max_point = max(overshoot_points, key=lambda point: point["value"])
        result.update(
            {
                "first_overshoot_time": first_point["time_label"],
                "first_overshoot_tc": first_point["tc_name"],
                "first_overshoot_tc_number": first_point["tc_number"],
                "first_overshoot_value": first_point["value"],
                "max_overshoot_time": max_point["time_label"],
                "max_overshoot_tc": max_point["tc_name"],
                "max_overshoot_tc_number": max_point["tc_number"],
                "max_overshoot_value": max_point["value"],
            }
        )

    return result


def analyze_interval_overshoot(
    corrected_df: pd.DataFrame, selected_row_indices, setpoint: float, tolerance: float
):
    time_col = corrected_df.columns[0]
    tc_columns = list(corrected_df.columns[1:])
    allowed_max = setpoint + tolerance
    overshoot_points = []

    for local_index, row_index in enumerate(selected_row_indices):
        row = corrected_df.iloc[row_index]
        time_label = format_axis_value(row[time_col])

        for tc_name in tc_columns:
            value = float(row[tc_name])
            if value > allowed_max:
                overshoot_points.append(
                    {
                        "time_label": time_label,
                        "tc_name": tc_name,
                        "tc_number": get_tc_number_from_name(tc_name),
                        "value": value,
                        "row_index": row_index,
                        "window_index": local_index,
                    }
                )

    result = {
        "overshoot_result": "VAR" if overshoot_points else "YOK",
        "overshoot_point_count": len(overshoot_points),
    }

    if overshoot_points:
        first_point = overshoot_points[0]
        max_point = max(overshoot_points, key=lambda point: point["value"])
        result.update(
            {
                "first_overshoot_time": first_point["time_label"],
                "first_overshoot_tc": first_point["tc_name"],
                "first_overshoot_tc_number": first_point["tc_number"],
                "first_overshoot_value": first_point["value"],
                "first_overshoot_row_index": first_point["row_index"],
                "first_overshoot_window_index": first_point["window_index"],
                "max_overshoot_time": max_point["time_label"],
                "max_overshoot_tc": max_point["tc_name"],
                "max_overshoot_tc_number": max_point["tc_number"],
                "max_overshoot_value": max_point["value"],
                "max_overshoot_row_index": max_point["row_index"],
                "max_overshoot_window_index": max_point["window_index"],
            }
        )

    return result


def evaluate_interval(
    corrected_df: pd.DataFrame,
    interval_correction,
):
    interval_config = interval_correction["config"]
    tc_cf_map = interval_correction["tc_cf_map"]
    logger_cf_map = interval_correction["logger_cf_map"]
    interval_df, invalid_rows, selected_row_indices = filter_corrected_data_by_interval(
        corrected_df,
        interval_config["start_time"],
        interval_config["end_time"],
    )
    summary_df, all_corrected_values = summarize_corrected_data(
        interval_df, tc_cf_map, logger_cf_map
    )
    summary_df, result_info = evaluate(
        summary_df,
        all_corrected_values,
        interval_config["setpoint"],
        interval_config["tolerance"],
    )
    result_info.update(find_extreme_points_in_rows(corrected_df, selected_row_indices))
    result_info.update(
        analyze_interval_overshoot(
            corrected_df,
            selected_row_indices,
            interval_config["setpoint"],
            interval_config["tolerance"],
        )
    )
    result_info["interval_label"] = interval_config["label"]
    result_info["interval_index"] = interval_config["index"]
    result_info["interval_start_time"] = interval_config["start_time"]
    result_info["interval_end_time"] = interval_config["end_time"]
    result_info["setpoint"] = interval_config["setpoint"]
    result_info["tolerance"] = interval_config["tolerance"]
    result_info["tc_cf_column"] = interval_correction["tc_cf_column"]
    result_info["logger_cf_column"] = interval_correction["logger_cf_column"]
    result_info["evaluated_row_count"] = len(interval_df)
    result_info["skipped_time_row_count"] = len(invalid_rows)
    result_info["overall_result_reason"] = ""

    if result_info["overshoot_result"] == "VAR":
        result_info["overall_result"] = "FAIL"
        result_info["overall_result_reason"] = (
            "Bu aralikta overshoot tespit edildigi icin aralik sonucu FAIL."
        )

    return {
        "config": interval_config,
        "interval_df": interval_df,
        "summary_df": summary_df,
        "result_info": result_info,
        "invalid_rows": invalid_rows,
        "selected_row_indices": selected_row_indices,
        "tc_cf_column": interval_correction["tc_cf_column"],
        "logger_cf_column": interval_correction["logger_cf_column"],
    }


def build_overall_summary(interval_results, tc_count: int):
    summary_rows = []
    overall_result = "PASS"
    failed_labels = []

    for interval_result in interval_results:
        config = interval_result["config"]
        info = interval_result["result_info"]
        summary_rows.append(
            {
                "INTERVAL": config["label"],
                "TIME_RANGE": format_time_range(config["start_time"], config["end_time"]),
                "SETPOINT": round_to_one_decimal(config["setpoint"]),
                "TOLERANCE": round_to_one_decimal(config["tolerance"]),
                "TC_CF_COLUMN": info["tc_cf_column"],
                "LOGGER_CF_COLUMN": info["logger_cf_column"],
                "RESULT": info["overall_result"],
                "OVERSHOOT": info["overshoot_result"],
                "HOTTEST_TC": info["hottest_tc"],
                "HOTTEST_VALUE": info["hottest_value"],
                "COLDEST_TC": info["coldest_tc"],
                "COLDEST_VALUE": info["coldest_value"],
                "SPREAD": info["spread"],
                "ROW_COUNT": info["evaluated_row_count"],
                "THERMOCOUPLE_COUNT": tc_count,
            }
        )
        if info["overall_result"] == "FAIL":
            overall_result = "FAIL"
            failed_labels.append(config["label"])

    overall_df = pd.DataFrame(summary_rows)
    return overall_df, overall_result, failed_labels


def create_report(
    summary_df,
    result_info,
    setpoint,
    tolerance,
    evaluation_start_time,
    stabilization_info,
    skipped_time_rows,
):
    lines = []
    hottest_tc_number = result_info.get("hottest_tc_number")
    coldest_tc_number = result_info.get("coldest_tc_number")

    lines.append("=" * 100)
    lines.append("AMS 2750 TUS HESAP RAPORU")
    lines.append("=" * 100)
    lines.append(f"Set Noktasi           : {format_number(setpoint)}")
    lines.append(f"Tolerans              : +/-{format_number(tolerance)}")
    lines.append(
        f"Degerlendirme Baslangici: {format_time_for_display(evaluation_start_time)}"
    )
    lines.append(
        f"Degerlendirme Araligi : {format_evaluation_window(evaluation_start_time)}"
    )
    lines.append(f"Degerlendirilen Veri Sayisi: {result_info['evaluated_row_count']}")
    lines.append(f"Atlanan Saat Satiri    : {len(skipped_time_rows)}")
    lines.append(
        f"Stabilizasyon Penceresi: {format_minutes_for_display(result_info['stabilization_window_minutes'])}"
    )
    lines.append(f"Izin Verilen Alt Limit: {format_number(result_info['allowed_min'])}")
    lines.append(f"Izin Verilen Ust Limit: {format_number(result_info['allowed_max'])}")
    lines.append("-" * 100)

    for i, row in summary_df.iterrows():
        tc_number = get_tc_number_from_name(row["TC"])
        lines.append(f"TC No                 : {tc_number if tc_number is not None else i + 1}")
        lines.append(f"TC Adi                : {row['TC']}")
        lines.append(f"Ham Min               : {format_number(row['RAW_MIN'])}")
        lines.append(f"Ham Max               : {format_number(row['RAW_MAX'])}")
        lines.append(f"TC Correction Factor  : {format_signed_number(row['TC_CF'])}")
        lines.append(f"Datalogger CF         : {format_signed_number(row['DATALOGGER_CF'])}")
        lines.append(f"Toplam CF             : {format_signed_number(row['TOTAL_CF'])}")
        lines.append(f"Min Duzeltilmis       : {format_number(row['MIN_CORRECTED'])}")
        lines.append(f"Max Duzeltilmis       : {format_number(row['MAX_CORRECTED'])}")
        lines.append(f"Ort. Duzeltilmis      : {format_number(row['AVG_CORRECTED'])}")
        lines.append(f"TC Sonucu             : {row['POINT_RESULT']}")
        lines.append("-" * 100)

    lines.append("DEGERLENDIRME ARALIGI SONUCLARI")
    lines.append("-" * 100)
    lines.append(f"Aralik Minimum        : {format_number(result_info['overall_min'])}")
    lines.append(f"Aralik Maximum        : {format_number(result_info['overall_max'])}")
    lines.append(f"Uniformity Spread     : {format_number(result_info['spread'])}")
    if coldest_tc_number is not None:
        lines.append(
            f"En Soguk TC No        : {coldest_tc_number} ({result_info['coldest_tc']})"
        )
    else:
        lines.append(f"En Soguk TC           : {result_info['coldest_tc']}")
    lines.append(f"En Soguk Deger        : {format_number(result_info['coldest_value'])}")
    if result_info.get("coldest_time_label"):
        lines.append(f"En Soguk Zaman        : {result_info['coldest_time_label']}")
    if hottest_tc_number is not None:
        lines.append(
            f"En Sicak TC No        : {hottest_tc_number} ({result_info['hottest_tc']})"
        )
    else:
        lines.append(f"En Sicak TC           : {result_info['hottest_tc']}")
    lines.append(f"En Sicak Deger        : {format_number(result_info['hottest_value'])}")
    if result_info.get("hottest_time_label"):
        lines.append(f"En Sicak Zaman        : {result_info['hottest_time_label']}")
    if hottest_tc_number is not None and coldest_tc_number is not None:
        lines.append(
            f"Sonuc Ozeti           : Degerlendirme araligindaki sonuclara gore {hottest_tc_number} nolu TC en sicak, {coldest_tc_number} nolu TC en soguk"
        )
        lines.append(
            f"Aralik TC Ozeti       : {format_time_phrase(evaluation_start_time)} en sicak TC: {hottest_tc_number}, en soguk TC: {coldest_tc_number}"
        )
    else:
        lines.append(
            f"Sonuc Ozeti           : Degerlendirme araligindaki sonuclara gore en sicak TC {result_info['hottest_tc']}, en soguk TC {result_info['coldest_tc']}"
        )
        lines.append(
            f"Aralik TC Ozeti       : {format_time_phrase(evaluation_start_time)} en sicak TC: {result_info['hottest_tc']}, en soguk TC: {result_info['coldest_tc']}"
        )
    lines.append(f"Degerlendirme Sonucu  : {result_info['overall_result']}")
    if result_info.get("overall_result_reason"):
        lines.append(f"Degerlendirme Notu    : {result_info['overall_result_reason']}")
    if stabilization_info is None:
        lines.append("Stabilizasyon         : Yapilmadi")
    else:
        lines.append(f"Stabilizasyon         : {stabilization_info['status']}")
        if stabilization_info["status"] == "BULUNDU":
            lines.append(f"Stabil Baslangic      : {stabilization_info['start_label']}")
            lines.append(f"Stabil Bitis          : {stabilization_info['end_label']}")
            lines.append(f"Stabil Veri Sayisi    : {stabilization_info['row_count']}")
        else:
            lines.append(f"Stabilizasyon Notu    : {stabilization_info['message']}")
    lines.append("-" * 100)
    lines.append("TUM VERI OVERSHOOT KONTROLU")
    lines.append("-" * 100)
    lines.append(f"Overshoot Var Mi      : {result_info['overshoot_result']}")
    lines.append(f"Overshoot Nokta Sayisi: {result_info['overshoot_point_count']}")
    if result_info["overshoot_result"] == "VAR":
        first_tc_number = result_info.get("first_overshoot_tc_number")
        if first_tc_number is not None:
            lines.append(
                f"Ilk Overshoot TC      : {first_tc_number} ({result_info['first_overshoot_tc']})"
            )
        else:
            lines.append(
                f"Ilk Overshoot TC      : {result_info['first_overshoot_tc']}"
            )
        lines.append(
            f"Ilk Overshoot Zaman   : {result_info['first_overshoot_time']}"
        )
        lines.append(
            f"Ilk Overshoot Deger   : {format_number(result_info['first_overshoot_value'])}"
        )

        max_tc_number = result_info.get("max_overshoot_tc_number")
        if max_tc_number is not None:
            lines.append(
                f"Maksimum Overshoot TC : {max_tc_number} ({result_info['max_overshoot_tc']})"
            )
        else:
            lines.append(
                f"Maksimum Overshoot TC : {result_info['max_overshoot_tc']}"
            )
        lines.append(
            f"Maksimum Overshoot Zaman: {result_info['max_overshoot_time']}"
        )
        lines.append(
            f"Maksimum Overshoot Deger: {format_number(result_info['max_overshoot_value'])}"
        )
    lines.append("=" * 100)

    return "\n".join(lines)


def create_multi_interval_report(interval_results, overall_result, failed_labels, tc_count: int):
    lines = []
    lines.append("=" * 100)
    lines.append("AMS 2750 COKLU SET NOKTASI RAPORU")
    lines.append("=" * 100)
    lines.append(f"Toplam Aralik         : {len(interval_results)}")
    lines.append(f"Toplam Thermocouple   : {tc_count}")
    lines.append(f"Genel Sonuc           : {overall_result}")
    if failed_labels:
        lines.append(f"Basarisiz Araliklar   : {', '.join(failed_labels)}")
    else:
        lines.append("Basarisiz Araliklar   : Yok")

    for interval_result in interval_results:
        config = interval_result["config"]
        info = interval_result["result_info"]

        lines.append("-" * 100)
        lines.append(config["label"])
        lines.append("-" * 100)
        lines.append(
            f"Calisma Araligi       : {format_time_range(config['start_time'], config['end_time'])}"
        )
        lines.append(f"Set Noktasi           : {format_number(config['setpoint'])}")
        lines.append(f"Tolerans              : +/-{format_number(config['tolerance'])}")
        lines.append(f"TC CF Kolonu          : {info['tc_cf_column']}")
        lines.append(f"Logger CF Kolonu      : {info['logger_cf_column']}")
        lines.append(f"Izin Verilen Alt Limit: {format_number(info['allowed_min'])}")
        lines.append(f"Izin Verilen Ust Limit: {format_number(info['allowed_max'])}")
        lines.append(f"Degerlendirilen Veri  : {info['evaluated_row_count']}")
        lines.append(f"Atlanan Saat Satiri   : {info['skipped_time_row_count']}")
        lines.append(f"Aralik Minimum        : {format_number(info['overall_min'])}")
        lines.append(f"Aralik Maximum        : {format_number(info['overall_max'])}")
        lines.append(f"Uniformity Spread     : {format_number(info['spread'])}")

        hottest_tc_number = info.get("hottest_tc_number")
        coldest_tc_number = info.get("coldest_tc_number")
        if hottest_tc_number is not None:
            lines.append(
                f"En Sicak TC No        : {hottest_tc_number} ({info['hottest_tc']})"
            )
        else:
            lines.append(f"En Sicak TC           : {info['hottest_tc']}")
        lines.append(f"En Sicak Deger        : {format_number(info['hottest_value'])}")
        lines.append(f"En Sicak Zaman        : {info['hottest_time_label']}")

        if coldest_tc_number is not None:
            lines.append(
                f"En Soguk TC No        : {coldest_tc_number} ({info['coldest_tc']})"
            )
        else:
            lines.append(f"En Soguk TC           : {info['coldest_tc']}")
        lines.append(f"En Soguk Deger        : {format_number(info['coldest_value'])}")
        lines.append(f"En Soguk Zaman        : {info['coldest_time_label']}")

        if hottest_tc_number is not None and coldest_tc_number is not None:
            lines.append(
                f"Aralik TC Ozeti       : {config['start_time'].strftime('%H:%M')} - {config['end_time'].strftime('%H:%M')} arasinda en sicak TC: {hottest_tc_number}, en soguk TC: {coldest_tc_number}"
            )
        else:
            lines.append(
                f"Aralik TC Ozeti       : {config['start_time'].strftime('%H:%M')} - {config['end_time'].strftime('%H:%M')} arasinda en sicak TC: {info['hottest_tc']}, en soguk TC: {info['coldest_tc']}"
            )

        lines.append(f"Overshoot Var Mi      : {info['overshoot_result']}")
        lines.append(f"Overshoot Nokta Sayisi: {info['overshoot_point_count']}")
        if info["overshoot_result"] == "VAR":
            first_tc_number = info.get("first_overshoot_tc_number")
            if first_tc_number is not None:
                lines.append(
                    f"Ilk Overshoot TC      : {first_tc_number} ({info['first_overshoot_tc']})"
                )
            else:
                lines.append(f"Ilk Overshoot TC      : {info['first_overshoot_tc']}")
            lines.append(f"Ilk Overshoot Zaman   : {info['first_overshoot_time']}")
            lines.append(f"Ilk Overshoot Deger   : {format_number(info['first_overshoot_value'])}")

            max_tc_number = info.get("max_overshoot_tc_number")
            if max_tc_number is not None:
                lines.append(
                    f"Maksimum Overshoot TC : {max_tc_number} ({info['max_overshoot_tc']})"
                )
            else:
                lines.append(f"Maksimum Overshoot TC : {info['max_overshoot_tc']}")
            lines.append(f"Maksimum Overshoot Zaman: {info['max_overshoot_time']}")
            lines.append(f"Maksimum Overshoot Deger: {format_number(info['max_overshoot_value'])}")

        lines.append(f"Aralik Sonucu         : {info['overall_result']}")
        if info.get("overall_result_reason"):
            lines.append(f"Aralik Notu           : {info['overall_result_reason']}")

    lines.append("=" * 100)
    return "\n".join(lines)


def create_interval_chart(interval_result, chart_path: Path):
    try:
        import matplotlib.pyplot as plt
    except ImportError as exc:
        raise ImportError(
            "Grafik olusturmak icin matplotlib gerekli. Kurmak icin: py -m pip install matplotlib"
        ) from exc

    interval_df = interval_result["interval_df"]
    config = interval_result["config"]
    info = interval_result["result_info"]
    time_col = interval_df.columns[0]
    tc_columns = list(interval_df.columns[1:])
    x_positions = list(range(len(interval_df)))
    x_labels = [format_axis_value(value) for value in interval_df[time_col].tolist()]

    fig, ax = plt.subplots(figsize=(15, 8))

    for tc_name in tc_columns:
        ax.plot(x_positions, interval_df[tc_name], linewidth=1.5, alpha=0.7, label=tc_name)

    ax.axhline(config["setpoint"], color="black", linewidth=1.4, label="Set Noktasi")
    ax.axhline(
        config["setpoint"] - config["tolerance"],
        color="red",
        linestyle="--",
        linewidth=1.2,
        label="Alt Limit",
    )
    ax.axhline(
        config["setpoint"] + config["tolerance"],
        color="green",
        linestyle="--",
        linewidth=1.2,
        label="Ust Limit",
    )

    hottest_index = info.get("hottest_window_index")
    coldest_index = info.get("coldest_window_index")
    if hottest_index is not None:
        ax.scatter(
            hottest_index,
            info["hottest_value"],
            color="orangered",
            s=90,
            edgecolors="white",
            linewidths=1.0,
            zorder=6,
            label=f"En Sicak Nokta - {info['hottest_tc']}",
        )
        ax.annotate(
            f"En sicak: {info['hottest_tc']} ({format_number(info['hottest_value'])})",
            xy=(hottest_index, info["hottest_value"]),
            xytext=(10, 10),
            textcoords="offset points",
            fontsize=9,
            color="orangered",
            bbox={"boxstyle": "round,pad=0.25", "fc": "white", "ec": "orangered", "alpha": 0.9},
        )

    if coldest_index is not None:
        ax.scatter(
            coldest_index,
            info["coldest_value"],
            color="royalblue",
            s=90,
            edgecolors="white",
            linewidths=1.0,
            zorder=6,
            label=f"En Soguk Nokta - {info['coldest_tc']}",
        )
        ax.annotate(
            f"En soguk: {info['coldest_tc']} ({format_number(info['coldest_value'])})",
            xy=(coldest_index, info["coldest_value"]),
            xytext=(10, -18),
            textcoords="offset points",
            fontsize=9,
            color="royalblue",
            bbox={"boxstyle": "round,pad=0.25", "fc": "white", "ec": "royalblue", "alpha": 0.9},
        )

    ax.set_title(
        f"{config['label']} - {format_time_range(config['start_time'], config['end_time'])}",
        fontsize=15,
    )
    ax.set_xlabel(str(time_col))
    ax.set_ylabel("Duzeltilmis Sicaklik")
    ax.grid(True, linestyle=":", alpha=0.5)

    if x_positions:
        max_ticks = 12
        step = max(1, math.ceil(len(x_positions) / max_ticks))
        tick_positions = x_positions[::step]
        if tick_positions[-1] != x_positions[-1]:
            tick_positions.append(x_positions[-1])
        tick_labels = [x_labels[position] for position in tick_positions]
        ax.set_xticks(tick_positions)
        ax.set_xticklabels(tick_labels)

    ax.legend(loc="center left", bbox_to_anchor=(1.02, 0.5), ncol=1, frameon=True)
    plt.xticks(rotation=45, ha="right")
    fig.tight_layout(rect=(0, 0, 0.82, 1))
    fig.savefig(chart_path, dpi=200, bbox_inches="tight")
    plt.close(fig)


def create_all_intervals_chart(raw_df, corrected_df, interval_results, chart_path: Path):
    try:
        import matplotlib.pyplot as plt
    except ImportError as exc:
        raise ImportError(
            "Grafik olusturmak icin matplotlib gerekli. Kurmak icin: py -m pip install matplotlib"
        ) from exc

    chart_df = build_chart_display_data(raw_df, corrected_df)
    time_col = chart_df.columns[0]
    tc_columns = list(chart_df.columns[1:])
    x_positions = list(range(len(chart_df)))
    x_labels = [format_axis_value(value) for value in chart_df[time_col].tolist()]
    interval_colors = ["#FFF0C9", "#DDF2FF", "#E4FFD9"]

    fig, ax = plt.subplots(figsize=(16, 9))

    for tc_name in tc_columns:
        ax.plot(x_positions, chart_df[tc_name], linewidth=1.2, alpha=0.35, label=tc_name)

    for index, interval_result in enumerate(interval_results):
        config = interval_result["config"]
        info = interval_result["result_info"]
        row_indices = interval_result["selected_row_indices"]
        if not row_indices:
            continue

        start_index = row_indices[0]
        end_index = row_indices[-1]
        interval_color = interval_colors[index % len(interval_colors)]

        ax.axvspan(
            start_index - 0.5,
            end_index + 0.5,
            color=interval_color,
            alpha=0.25,
            label=f"{config['label']} Araligi",
        )
        ax.hlines(
            config["setpoint"],
            start_index,
            end_index,
            colors="black",
            linestyles="-",
            linewidth=1.1,
        )
        ax.hlines(
            config["setpoint"] - config["tolerance"],
            start_index,
            end_index,
            colors="red",
            linestyles="--",
            linewidth=1.0,
        )
        ax.hlines(
            config["setpoint"] + config["tolerance"],
            start_index,
            end_index,
            colors="green",
            linestyles="--",
            linewidth=1.0,
        )

        hottest_row_index = info.get("hottest_row_index")
        coldest_row_index = info.get("coldest_row_index")
        if hottest_row_index is not None:
            ax.scatter(
                hottest_row_index,
                info["hottest_value"],
                color="orangered",
                s=70,
                edgecolors="white",
                linewidths=0.9,
                zorder=6,
            )
            ax.annotate(
                f"{config['label']} sicak",
                xy=(hottest_row_index, info["hottest_value"]),
                xytext=(8, 10),
                textcoords="offset points",
                fontsize=8,
                color="orangered",
                bbox={"boxstyle": "round,pad=0.2", "fc": "white", "ec": "orangered", "alpha": 0.85},
            )

        if coldest_row_index is not None:
            ax.scatter(
                coldest_row_index,
                info["coldest_value"],
                color="royalblue",
                s=70,
                edgecolors="white",
                linewidths=0.9,
                zorder=6,
            )
            ax.annotate(
                f"{config['label']} soguk",
                xy=(coldest_row_index, info["coldest_value"]),
                xytext=(8, -18),
                textcoords="offset points",
                fontsize=8,
                color="royalblue",
                bbox={"boxstyle": "round,pad=0.2", "fc": "white", "ec": "royalblue", "alpha": 0.85},
            )

    ax.set_title("AMS 2750 TUS - Tum Araliklar", fontsize=15)
    ax.set_xlabel(str(time_col))
    ax.set_ylabel("Sicaklik")
    ax.grid(True, linestyle=":", alpha=0.5)

    if x_positions:
        max_ticks = 12
        step = max(1, math.ceil(len(x_positions) / max_ticks))
        tick_positions = x_positions[::step]
        if tick_positions[-1] != x_positions[-1]:
            tick_positions.append(x_positions[-1])
        tick_labels = [x_labels[position] for position in tick_positions]
        ax.set_xticks(tick_positions)
        ax.set_xticklabels(tick_labels)

    ax.legend(loc="center left", bbox_to_anchor=(1.02, 0.5), ncol=1, frameon=True)
    plt.xticks(rotation=45, ha="right")
    fig.tight_layout(rect=(0, 0, 0.82, 1))
    fig.savefig(chart_path, dpi=200, bbox_inches="tight")
    plt.close(fig)


def save_multi_interval_outputs(
    raw_df,
    corrected_df,
    interval_results,
    overall_summary_df,
    report_text,
    excel_path: Path,
    report_path: Path,
    full_chart_path: Path,
    interval_chart_paths,
):
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        corrected_df.to_excel(writer, sheet_name="Corrected_Data", index=False)
        overall_summary_df.to_excel(writer, sheet_name="Overall_Summary", index=False)

        for interval_result in interval_results:
            interval_index = interval_result["config"]["index"]
            interval_result["interval_df"].to_excel(
                writer,
                sheet_name=f"Interval_{interval_index}_Data",
                index=False,
            )
            interval_result["summary_df"].to_excel(
                writer,
                sheet_name=f"Interval_{interval_index}_Summary",
                index=False,
            )

    report_path.write_text(report_text, encoding="utf-8")
    create_all_intervals_chart(raw_df, corrected_df, interval_results, full_chart_path)

    for interval_result, chart_path in zip(interval_results, interval_chart_paths):
        create_interval_chart(interval_result, chart_path)

    print("\nDosyalar olusturuldu:")
    print(f"- {excel_path}")
    print(f"- {report_path}")
    print(f"- {full_chart_path}")
    for chart_path in interval_chart_paths:
        print(f"- {chart_path}")


def create_rise_chart(
    corrected_df,
    setpoint,
    tolerance,
    chart_path: Path,
    chart_title: str,
    evaluation_start_time: time | None = None,
    hottest_point: dict | None = None,
    coldest_point: dict | None = None,
    use_window_indices: bool = False,
    annotate_extremes: bool = False,
):
    try:
        import matplotlib.pyplot as plt
    except ImportError as exc:
        raise ImportError(
            "Grafik olusturmak icin matplotlib gerekli. Kurmak icin: py -m pip install matplotlib"
        ) from exc

    time_col = corrected_df.columns[0]
    tc_columns = list(corrected_df.columns[1:])
    raw_x_values = corrected_df[time_col].tolist()
    x_positions = list(range(len(raw_x_values)))
    x_labels = [format_axis_value(value) for value in raw_x_values]
    if use_window_indices:
        highlight_mask = [True] * len(corrected_df)
    else:
        highlight_mask, _ = build_evaluation_row_mask(corrected_df, evaluation_start_time)
    highlight_indices = [index for index, keep in enumerate(highlight_mask) if keep]
    allowed_min = setpoint - tolerance
    allowed_max = setpoint + tolerance

    fig, ax = plt.subplots(figsize=(16, 9))

    for tc_name in tc_columns:
        ax.plot(
            x_positions,
            corrected_df[tc_name],
            linewidth=1.2,
            alpha=0.35 if highlight_indices else 0.65,
            label=tc_name,
        )

    if not use_window_indices and highlight_indices:
        first_highlight_index = highlight_indices[0]
        if first_highlight_index > 0:
            ax.axvspan(
                -0.5,
                first_highlight_index - 0.5,
                color="lightgray",
                alpha=0.25,
                label="Degerlendirme Disi Bolge",
            )

    def plot_highlight_segment(tc_name: str | None, color: str, label: str, zorder: int):
        if tc_name not in tc_columns or not highlight_indices:
            return
        x_segment = highlight_indices
        y_segment = [float(corrected_df[tc_name].iloc[index]) for index in x_segment]
        ax.plot(
            x_segment,
            y_segment,
            linewidth=3.2,
            color=color,
            alpha=0.95,
            label=label,
            zorder=zorder,
        )

    hottest_tc_name = hottest_point.get("hottest_tc") if hottest_point else None
    coldest_tc_name = coldest_point.get("coldest_tc") if coldest_point else None
    plot_highlight_segment(coldest_tc_name, "royalblue", f"En Soguk TC - {coldest_tc_name}", 4)
    plot_highlight_segment(hottest_tc_name, "orangered", f"En Sicak TC - {hottest_tc_name}", 5)

    ax.axhline(setpoint, color="black", linestyle="-", linewidth=1.4, label="Set Noktasi")
    ax.axhline(
        allowed_min,
        color="red",
        linestyle="--",
        linewidth=1.2,
        label="Alt Limit",
    )
    ax.axhline(
        allowed_max,
        color="green",
        linestyle="--",
        linewidth=1.2,
        label="Ust Limit",
    )

    if evaluation_start_time is not None:
        for index, value in enumerate(raw_x_values):
            row_time = extract_time_of_day(value)
            if row_time is not None and row_time >= evaluation_start_time:
                ax.axvline(
                    index,
                    color="navy",
                    linestyle=":",
                    linewidth=1.4,
                    label=f"Degerlendirme Baslangici ({format_time_for_display(evaluation_start_time)})",
                )
                break

    if annotate_extremes and x_positions:
        if hottest_point and hottest_point.get("hottest_tc") in tc_columns:
            hottest_index_key = "hottest_window_index" if use_window_indices else "hottest_row_index"
            hottest_index = hottest_point.get(hottest_index_key)
            hottest_value = hottest_point.get("hottest_value")
            hottest_tc = hottest_point.get("hottest_tc")
            if hottest_index is not None and hottest_value is not None and 0 <= hottest_index < len(x_positions):
                ax.scatter(
                    hottest_index,
                    hottest_value,
                    color="orangered",
                    s=85,
                    edgecolors="white",
                    linewidths=1.0,
                    zorder=6,
                    label=f"En Sicak Nokta - {hottest_tc}",
                )
                ax.annotate(
                    f"En sicak: {hottest_tc} ({format_number(hottest_value)})",
                    xy=(hottest_index, hottest_value),
                    xytext=(10, 10),
                    textcoords="offset points",
                    color="orangered",
                    fontsize=9,
                    bbox={"boxstyle": "round,pad=0.25", "fc": "white", "ec": "orangered", "alpha": 0.9},
                )

        if coldest_point and coldest_point.get("coldest_tc") in tc_columns:
            coldest_index_key = "coldest_window_index" if use_window_indices else "coldest_row_index"
            coldest_index = coldest_point.get(coldest_index_key)
            coldest_value = coldest_point.get("coldest_value")
            coldest_tc = coldest_point.get("coldest_tc")
            if coldest_index is not None and coldest_value is not None and 0 <= coldest_index < len(x_positions):
                ax.scatter(
                    coldest_index,
                    coldest_value,
                    color="royalblue",
                    s=85,
                    edgecolors="white",
                    linewidths=1.0,
                    zorder=6,
                    label=f"En Soguk Nokta - {coldest_tc}",
                )
                ax.annotate(
                    f"En soguk: {coldest_tc} ({format_number(coldest_value)})",
                    xy=(coldest_index, coldest_value),
                    xytext=(10, -18),
                    textcoords="offset points",
                    color="royalblue",
                    fontsize=9,
                    bbox={"boxstyle": "round,pad=0.25", "fc": "white", "ec": "royalblue", "alpha": 0.9},
                )

    ax.set_title(chart_title, fontsize=15)
    ax.set_xlabel(str(time_col))
    ax.set_ylabel("Duzeltilmis Sicaklik")
    ax.grid(True, linestyle=":", alpha=0.5)

    if x_positions:
        max_ticks = 12
        step = max(1, math.ceil(len(x_positions) / max_ticks))
        tick_positions = x_positions[::step]
        if tick_positions[-1] != x_positions[-1]:
            tick_positions.append(x_positions[-1])
        tick_labels = [x_labels[pos] for pos in tick_positions]
        ax.set_xticks(tick_positions)
        ax.set_xticklabels(tick_labels)

    legend_columns = 1 if len(tc_columns) <= 12 else 2
    ax.legend(loc="center left", bbox_to_anchor=(1.02, 0.5), ncol=legend_columns, frameon=True)

    plt.xticks(rotation=45, ha="right")
    fig.tight_layout(rect=(0, 0, 0.82, 1))
    fig.savefig(chart_path, dpi=200, bbox_inches="tight")
    plt.close(fig)


def save_outputs(
    corrected_df,
    evaluation_df,
    summary_df,
    report_text,
    excel_path: Path,
    report_path: Path,
    full_chart_path: Path,
    evaluation_chart_path: Path,
    setpoint: float,
    tolerance: float,
    evaluation_start_time: time | None,
    result_info,
):
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        corrected_df.to_excel(writer, sheet_name="Corrected_Data", index=False)
        evaluation_df.to_excel(writer, sheet_name="Evaluation_Data", index=False)
        summary_df.to_excel(writer, sheet_name="Summary", index=False)

    report_path.write_text(report_text, encoding="utf-8")
    create_rise_chart(
        corrected_df,
        setpoint,
        tolerance,
        full_chart_path,
        "AMS 2750 TUS Yukselme Grafigi - Tum Veri",
        evaluation_start_time,
        result_info,
        result_info,
        False,
        True,
    )
    create_rise_chart(
        evaluation_df,
        setpoint,
        tolerance,
        evaluation_chart_path,
        "AMS 2750 TUS Yukselme Grafigi - Degerlendirilen Veri",
        None,
        result_info,
        result_info,
        True,
        True,
    )

    print("\nDosyalar olusturuldu:")
    print(f"- {excel_path}")
    print(f"- {report_path}")
    print(f"- {full_chart_path}")
    print(f"- {evaluation_chart_path}")


def parse_args():
    parser = argparse.ArgumentParser(description="AMS 2750 TUS hesaplama araci")
    parser.add_argument("--raw-file", default=DEFAULT_RAW_FILE, help="Ham veri Excel dosyasi")
    parser.add_argument(
        "--tc-cf-file",
        default=DEFAULT_TC_CF_FILE,
        help="Thermocouple correction factor Excel dosyasi",
    )
    parser.add_argument(
        "--logger-cf-file",
        default=DEFAULT_LOGGER_CF_FILE,
        help="Datalogger correction factor Excel dosyasi",
    )
    parser.add_argument(
        "--interval",
        action="append",
        help="Aralik formati: baslangic|bitis|set noktasi|tolerans. Ornek: 11:32|11:50|60|2",
    )
    parser.add_argument(
        "--output-dir",
        help="Cikti dosyalarinin kaydedilecegi klasor. Verilmezse ham verinin klasoru kullanilir.",
    )
    parser.add_argument(
        "--non-interactive",
        action="store_true",
        help="Eksik arguman varsa soru sormak yerine hata verir.",
    )
    return parser.parse_args()


def resolve_inputs(args):
    raw_file = args.raw_file
    tc_cf_file = args.tc_cf_file
    logger_cf_file = args.logger_cf_file
    interval_args = args.interval or []

    if interval_args:
        if len(interval_args) != INTERVAL_COUNT:
            raise ValueError(f"--interval tam olarak {INTERVAL_COUNT} kez verilmelidir.")
        interval_configs = [
            parse_interval_arg(interval_text, interval_index + 1)
            for interval_index, interval_text in enumerate(interval_args)
        ]
        return raw_file, tc_cf_file, logger_cf_file, interval_configs

    if args.non_interactive:
        raise ValueError(
            f"--non-interactive kullaniminda --interval tam olarak {INTERVAL_COUNT} kez verilmelidir."
        )

    raw_file = ask_file_path("1) Ham veri Excel dosyasi", raw_file)
    tc_cf_file = ask_file_path("2) Thermocouple CF Excel dosyasi", tc_cf_file)
    logger_cf_file = ask_file_path("3) Datalogger CF Excel dosyasi", logger_cf_file)
    interval_configs = ask_interval_configs(INTERVAL_COUNT)
    return raw_file, tc_cf_file, logger_cf_file, interval_configs


def main():
    args = parse_args()

    print("=" * 100)
    print("AMS 2750 TUS HESAPLAMA PROGRAMI")
    print("=" * 100)

    try:
        raw_file, tc_cf_file, logger_cf_file, interval_configs = resolve_inputs(args)

        time_col, time_values, raw_tc_data = load_raw_data(raw_file)
        tc_count = len(raw_tc_data)
        tc_cf_data = load_cf_data(tc_cf_file, "Thermocouple CF")
        logger_cf_data = load_cf_data(logger_cf_file, "Datalogger CF")
        interval_corrections = prepare_interval_corrections(
            raw_tc_data, tc_cf_data, logger_cf_data, interval_configs
        )
        print(f"\nToplam {tc_count} adet thermocouple kolonu bulundu.")

        raw_df = build_raw_data_frame(time_col, time_values, raw_tc_data)
        corrected_df = build_combined_corrected_data(
            time_col, time_values, raw_tc_data, interval_corrections
        )

        interval_results = [
            evaluate_interval(corrected_df, interval_correction)
            for interval_correction in interval_corrections
        ]
        overall_summary_df, overall_result, failed_labels = build_overall_summary(
            interval_results, tc_count
        )

        report_text = create_multi_interval_report(
            interval_results, overall_result, failed_labels, tc_count
        )
        excel_path, report_path, full_chart_path, interval_chart_paths = make_output_paths(
            raw_file, args.output_dir
        )

        print("\n" + report_text)
        save_multi_interval_outputs(
            raw_df,
            corrected_df,
            interval_results,
            overall_summary_df,
            report_text,
            excel_path,
            report_path,
            full_chart_path,
            interval_chart_paths,
        )

    except Exception as exc:
        print(f"\nHata olustu: {exc}")
        raise SystemExit(1) from exc


if __name__ == "__main__":
    main()

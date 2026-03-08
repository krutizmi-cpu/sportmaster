from __future__ import annotations

import io
import math
import re
from pathlib import Path
from typing import Dict, Optional, Tuple

import pandas as pd
import streamlit as st
from difflib import SequenceMatcher
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

APP_DIR = Path(__file__).resolve().parent
DATA_DIR = APP_DIR / "data"
TEMPLATES_DIR = APP_DIR / "templates"
DEFAULT_COMMISSIONS_FILENAME = "sportmaster_commissions_2026-02-01.xlsx"

PRODUCT_REQUIRED_COLS = [
    "Артикул",
    "Наименование товара",
    "Себестоимость, ₽",
    "Цена продажи, ₽",
    "Вес факт, кг",
    "Длина, см",
    "Ширина, см",
    "Высота, см",
]

PRODUCT_OPTIONAL_COLS = [
    "Дней хранения FBSM",
    "Реклама, %",
    "Система налогообложения",
    "Налог, %",
    "Прочие расходы, ₽",
    "Целевая маржа, %",
    "Доля возвратов, %",
    "Доля невыкупа/отмен, %",
    "Тарифная группа",
    "Товарная группа 3 уровня",
]

COLUMN_ALIASES = {
    "sku": "Артикул",
    "артикул": "Артикул",
    "наименование": "Наименование товара",
    "наименование товара": "Наименование товара",
    "товар": "Наименование товара",
    "себестоимость": "Себестоимость, ₽",
    "себестоимость руб": "Себестоимость, ₽",
    "себестоимость ₽": "Себестоимость, ₽",
    "цена": "Цена продажи, ₽",
    "цена продажи": "Цена продажи, ₽",
    "цена продажи ₽": "Цена продажи, ₽",
    "price": "Цена продажи, ₽",
    "вес": "Вес факт, кг",
    "вес кг": "Вес факт, кг",
    "вес факт": "Вес факт, кг",
    "вес факт кг": "Вес факт, кг",
    "длина": "Длина, см",
    "длина см": "Длина, см",
    "ширина": "Ширина, см",
    "ширина см": "Ширина, см",
    "высота": "Высота, см",
    "высота см": "Высота, см",
    "дней хранения fbsm": "Дней хранения FBSM",
    "реклама": "Реклама, %",
    "реклама %": "Реклама, %",
    "система налогообложения": "Система налогообложения",
    "налог": "Налог, %",
    "налог %": "Налог, %",
    "прочие расходы": "Прочие расходы, ₽",
    "прочие расходы ₽": "Прочие расходы, ₽",
    "целевая маржа": "Целевая маржа, %",
    "целевая маржа %": "Целевая маржа, %",
    "доля возвратов": "Доля возвратов, %",
    "доля возвратов %": "Доля возвратов, %",
    "доля невыкупа отмен": "Доля невыкупа/отмен, %",
    "доля невыкупа отмен %": "Доля невыкупа/отмен, %",
    "тарифная группа": "Тарифная группа",
    "товарная группа 3 уровня": "Товарная группа 3 уровня",
}

TAX_SYSTEMS = {
    "ОСНО (22%)": 22.0,
    "УСН доходы (6%)": 6.0,
    "УСН доходы-расходы (15%)": 15.0,
    "Без налога (0%)": 0.0,
    "Своя ставка": None,
}

VISIBLE_COLUMNS = [
    "Артикул",
    "Наименование товара",
    "Схема",
    "Тарифная группа",
    "Товарная группа 3 уровня",
    "Как определили категорию",
    "Цена продажи, ₽",
    "Рекомендованная цена, ₽",
    "Себестоимость, ₽",
    "Комиссия, %",
    "Комиссия, ₽",
    "Логистика до покупателя, ₽",
    "Обратная логистика (ожидаемая), ₽",
    "Хранение FBSM, ₽",
    "Обработка брака, ₽",
    "Обработка излишков, ₽",
    "Реклама, ₽",
    "Налог, ₽",
    "Прочие расходы, ₽",
    "Полная себестоимость, ₽",
    "Выплата до себестоимости, ₽",
    "Прибыль, ₽",
    "Маржа к выручке, %",
    "Наценка на полную себестоимость, %",
    "Целевая маржа, %",
    "Прибыль при рекомендованной цене, ₽",
    "Маржа при рекомендованной цене, %",
    "Комментарий",
]

STOP_WORDS = {
    "для", "и", "с", "со", "на", "по", "под", "над", "из", "к", "ко", "от", "до", "или", "в", "во",
    "мужские", "мужской", "женские", "женский", "детские", "детский", "взрослые", "взрослый",
    "шт", "комплект", "набор", "унисекс", "товар", "спортивный", "спортивная", "спортивные",
}

TOKEN_SYNONYMS = {
    "велосипеды": "велосипед",
    "электровелосипеды": "электровелосипед",
    "кроссовки": "кроссовк",
    "ботинки": "ботинк",
    "перчатки": "перчатк",
    "варежки": "варежк",
    "мячи": "мяч",
    "гантели": "гантел",
    "вело": "велосипед",
    "горный": "горн",
    "беговые": "бегов",
    "лыжи": "лыж",
    "самокаты": "самокат",
    "велошлем": "шлем",
}


def normalize_text(value: str) -> str:
    s = str(value or "").strip().lower().replace("ё", "е")
    s = re.sub(r"[^0-9a-zа-я]+", " ", s)
    return " ".join(s.split())


def stem_token(token: str) -> str:
    t = normalize_text(token)
    if t in TOKEN_SYNONYMS:
        return TOKEN_SYNONYMS[t]
    for suffix in ["иями", "ями", "ами", "ями", "ов", "ев", "ей", "иях", "иях", "ия", "ья", "ье", "ий", "ый", "ой", "ая", "яя", "ое", "ее", "ые", "ие", "ам", "ям", "ах", "ях", "ом", "ем", "ую", "юю", "а", "я", "ы", "и", "е", "о", "у", "ю"]:
        if len(t) >= 6 and t.endswith(suffix):
            t = t[: -len(suffix)]
            break
    return TOKEN_SYNONYMS.get(t, t)


def text_tokens(value: str) -> set[str]:
    raw = normalize_text(value).split()
    tokens = {stem_token(tok) for tok in raw if tok and tok not in STOP_WORDS and len(tok) >= 3}
    return {t for t in tokens if t and t not in STOP_WORDS}


@st.cache_data(show_spinner=False)
def read_reference() -> pd.DataFrame:
    src = DATA_DIR / DEFAULT_COMMISSIONS_FILENAME
    if not src.exists():
        raise FileNotFoundError(
            f"Не найден файл data/{DEFAULT_COMMISSIONS_FILENAME}. Положите его в repo в папку data."
        )

    df = pd.read_excel(src)
    df.columns = [str(c).strip() for c in df.columns]
    expected = [
        "Товарная группа 1 уровня",
        "Товарная группа 2 уровня",
        "Товарная группа 3 уровня",
        "Тарифная группа",
        "Ставка комиссии FBSM, %",
        "Ставка комиссии FBS, %",
    ]
    missing = [c for c in expected if c not in df.columns]
    if missing:
        raise ValueError(f"В файле комиссий отсутствуют колонки: {', '.join(missing)}")

    for col in [
        "Товарная группа 1 уровня",
        "Товарная группа 2 уровня",
        "Товарная группа 3 уровня",
        "Тарифная группа",
    ]:
        df[col] = df[col].astype(str).str.strip()
        df[f"__norm_{col}"] = df[col].map(normalize_text)
        df[f"__tokens_{col}"] = df[col].map(text_tokens)

    combo_cols = ["Товарная группа 2 уровня", "Товарная группа 3 уровня", "Тарифная группа"]
    df["__search_text"] = df[combo_cols].fillna("").astype(str).agg(" ".join, axis=1)
    df["__tokens_combo"] = df["__search_text"].map(text_tokens)

    for col in ["Ставка комиссии FBSM, %", "Ставка комиссии FBS, %"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)

    return df


def prepare_products(df: pd.DataFrame) -> pd.DataFrame:
    renamed = {}
    for c in df.columns:
        renamed[c] = COLUMN_ALIASES.get(normalize_text(c), str(c).strip())
    df = df.rename(columns=renamed).copy()

    missing = [c for c in PRODUCT_REQUIRED_COLS if c not in df.columns]
    if missing:
        raise ValueError(f"В файле товаров не хватает колонок: {', '.join(missing)}")

    for c in PRODUCT_OPTIONAL_COLS:
        if c not in df.columns:
            df[c] = None

    numeric_cols = [
        "Себестоимость, ₽", "Цена продажи, ₽", "Вес факт, кг", "Длина, см", "Ширина, см", "Высота, см",
        "Дней хранения FBSM", "Реклама, %", "Налог, %", "Прочие расходы, ₽", "Целевая маржа, %",
        "Доля возвратов, %", "Доля невыкупа/отмен, %",
    ]
    for c in numeric_cols:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)

    df["Система налогообложения"] = df["Система налогообложения"].fillna("").astype(str).str.strip()
    return df


@st.cache_data(show_spinner=False)
def build_product_template_bytes() -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Товары"
    headers = PRODUCT_REQUIRED_COLS + PRODUCT_OPTIONAL_COLS
    ws.append(headers)
    sample_rows = [
        ["SM-001", "Кроссовки беговые мужские", 2500, 5990, 0.8, 32, 22, 12, 0, 5, "ОСНО (22%)", 0, 0, 20, 8, 3, "", ""],
        ["SM-002", "Велосипед горный", 18000, 32990, 14.5, 145, 25, 78, 75, 7, "УСН доходы (6%)", 0, 300, 18, 5, 2, "", ""],
    ]
    for row in sample_rows:
        ws.append(row)

    header_fill = PatternFill("solid", fgColor="1F4E78")
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for idx, col in enumerate(headers, start=1):
        ws.column_dimensions[get_column_letter(idx)].width = max(14, min(34, len(col) + 4))

    guide = wb.create_sheet("Описание полей")
    guide.append(["Колонка", "Описание"])
    explanations = [
        ("Артикул", "Обязательное поле."),
        ("Наименование товара", "Обязательное поле. Категория и тарифная группа определяются автоматически по названию."),
        ("Себестоимость, ₽", "Обязательное поле."),
        ("Цена продажи, ₽", "Обязательное поле."),
        ("Вес факт, кг", "Обязательное поле."),
        ("Длина, см / Ширина, см / Высота, см", "Обязательные поля для логистики."),
        ("Тарифная группа / Товарная группа 3 уровня", "Необязательно. Нужны только если хотите вручную переопределить автоподбор."),
        ("Дней хранения FBSM", "Нужно только для FBSM."),
        ("Реклама, %", "Процент от цены продажи."),
        ("Система налогообложения", "Можно оставить пустым и задать общий режим слева в приложении."),
        ("Налог, %", "Если заполнено, имеет приоритет над режимом налогообложения."),
        ("Прочие расходы, ₽", "Постоянные дополнительные расходы на единицу."),
        ("Целевая маржа, %", "Для рекомендованной цены."),
        ("Доля возвратов, % / Доля невыкупа/отмен, %", "Используются для ожидаемой обратной логистики FBSM."),
    ]
    for row in explanations:
        guide.append(row)
    for cell in guide[1]:
        cell.fill = header_fill
        cell.font = Font(color="FFFFFF", bold=True)
    guide.column_dimensions["A"].width = 38
    guide.column_dimensions["B"].width = 110

    stream = io.BytesIO()
    wb.save(stream)
    return stream.getvalue()


def normalize_commission_rate(value: float) -> float:
    v = float(value or 0.0)
    return v * 100 if v <= 1 else v


def commission_rate_for_tariff_group(reference_df: pd.DataFrame, tariff_group: str, scheme: str) -> Optional[float]:
    if not tariff_group:
        return None
    col = "Ставка комиссии FBSM, %" if scheme == "FBSM" else "Ставка комиссии FBS, %"
    rows = reference_df[reference_df["Тарифная группа"] == tariff_group]
    if rows.empty:
        return None
    return normalize_commission_rate(rows.iloc[0][col])


def resolve_tax_rate(row_tax_pct: float, row_tax_system: str, default_tax_system: str, manual_default_tax_pct: float) -> Tuple[float, str]:
    row_tax_system = str(row_tax_system or "").strip()
    if row_tax_pct and float(row_tax_pct) > 0:
        return float(row_tax_pct), "Ставка из строки товара"
    if row_tax_system:
        if row_tax_system in TAX_SYSTEMS and TAX_SYSTEMS[row_tax_system] is not None:
            return float(TAX_SYSTEMS[row_tax_system]), row_tax_system
        try:
            return float(str(row_tax_system).replace("%", "").replace(",", ".")), "Ставка из строки товара"
        except ValueError:
            pass
    if default_tax_system == "Своя ставка":
        return float(manual_default_tax_pct), f"Своя ставка ({manual_default_tax_pct:.1f}%)"
    return float(TAX_SYSTEMS.get(default_tax_system, 0.0) or 0.0), default_tax_system


def score_reference_match(name: str, row: pd.Series) -> float:
    name_norm = normalize_text(name)
    name_tokens = text_tokens(name)
    if not name_tokens:
        return 0.0

    scores = []
    for col in ["Товарная группа 3 уровня", "Тарифная группа", "Товарная группа 2 уровня"]:
        ref_norm = row[f"__norm_{col}"]
        ref_tokens = row[f"__tokens_{col}"]
        seq_score = SequenceMatcher(None, name_norm, ref_norm).ratio()
        overlap = len(name_tokens & ref_tokens) / max(len(ref_tokens), 1)
        name_cover = len(name_tokens & ref_tokens) / max(len(name_tokens), 1)
        contains = 1.0 if ref_norm and ref_norm in name_norm else 0.0
        scores.append(max(seq_score * 0.55 + overlap * 0.45, overlap * 0.85 + name_cover * 0.15, contains))

    combo_tokens = row["__tokens_combo"]
    combo_overlap = len(name_tokens & combo_tokens) / max(len(name_tokens), 1)
    has_core_object = any(tok in combo_tokens for tok in name_tokens)
    score = max(scores)
    score = max(score, combo_overlap * 0.9 + (0.08 if has_core_object else 0.0))

    if "велосипед" in name_tokens and "велосипед" in combo_tokens:
        score = max(score, 0.94)
    if "электровелосипед" in name_tokens and "электровелосипед" in combo_tokens:
        score = max(score, 0.96)
    return min(score, 0.99)


def resolve_tariff_group(reference_df: pd.DataFrame, product_name: str, manual_tariff_group: str, manual_group3: str) -> Tuple[str, str, Optional[float], str]:
    name_norm = normalize_text(product_name)
    manual_tariff_group = str(manual_tariff_group or "").strip()
    manual_group3 = str(manual_group3 or "").strip()

    if manual_tariff_group:
        rows = reference_df[reference_df["Тарифная группа"] == manual_tariff_group]
        if not rows.empty:
            group3 = str(rows.iloc[0]["Товарная группа 3 уровня"])
            return manual_tariff_group, group3, 1.0, "Тарифная группа из файла товаров"

    if manual_group3:
        rows = reference_df[reference_df["Товарная группа 3 уровня"] == manual_group3]
        if not rows.empty:
            tariff_group = str(rows.iloc[0]["Тарифная группа"])
            return tariff_group, manual_group3, 1.0, "Товарная группа 3 уровня из файла товаров"

    for col, note in [
        ("Товарная группа 3 уровня", "Точное совпадение с ТГ3"),
        ("Тарифная группа", "Точное совпадение с тарифной группой"),
        ("Товарная группа 2 уровня", "Точное совпадение со 2 уровнем"),
    ]:
        rows = reference_df[reference_df[f"__norm_{col}"] == name_norm]
        if not rows.empty:
            row = rows.iloc[0]
            return str(row["Тарифная группа"]), str(row["Товарная группа 3 уровня"]), 1.0, note

    name_tokens = text_tokens(product_name)
    if not name_tokens:
        return "", "", None, "Не найдено"

    candidates = []
    for _, ref_row in reference_df.iterrows():
        score = score_reference_match(product_name, ref_row)
        if score > 0:
            candidates.append((score, str(ref_row["Тарифная группа"]), str(ref_row["Товарная группа 3 уровня"])))

    if not candidates:
        return "", "", None, "Не найдено"

    candidates.sort(key=lambda x: x[0], reverse=True)
    best_score, best_tariff, best_lvl3 = candidates[0]
    if best_score >= 0.52:
        return best_tariff, best_lvl3, best_score, f"Автоподбор по названию ({best_score:.2f})"
    return "", "", best_score, "Не найдено"


def ceil_to(value: float, step: float) -> float:
    return 0.0 if value <= 0 else math.ceil(value / step) * step


def calc_fbs_billable_weight(weight_kg: float, length_cm: float, width_cm: float, height_cm: float, basis: str) -> float:
    volume = max(length_cm, 0) * max(width_cm, 0) * max(height_cm, 0)
    actual = max(weight_kg, 0)
    pr = max(actual, volume / 5000.0)
    cdek = max(actual, volume / 4000.0)
    if basis == "Почта России":
        chosen = pr
    elif basis == "СДЭК":
        chosen = cdek
    elif basis == "Фактический":
        chosen = actual
    else:
        chosen = max(pr, cdek)
    return float(math.ceil(chosen))


def calc_fbs_delivery(weight_kg: float, profile: str) -> float:
    if weight_kg <= 0:
        return 0.0
    base, extra = (220.0, 90.0) if profile == "220 + 90" else (200.0, 70.0)
    return base if weight_kg <= 2 else base + (weight_kg - 2) * extra


def calc_fbsm_delivery(weight_kg: float, profile: str) -> float:
    w = ceil_to(weight_kg, 0.1)
    if w <= 0:
        return 0.0
    up05, up10, extra = (56.0, 90.0, 60.0) if profile == "56 / 90 / 60" else (35.0, 75.0, 35.0)
    if w <= 0.5:
        return up05
    if w <= 1.0:
        return up10
    return up10 + (w - 1.0) * extra


def calc_fbsm_storage(weight_kg: float, storage_days: float) -> float:
    days = max(int(storage_days), 0)
    if days <= 60:
        return 0.0
    w = ceil_to(weight_kg, 0.1)
    days_61_90 = max(min(days, 90) - 60, 0)
    days_91_plus = max(days - 90, 0)
    return w * 3 * days_61_90 + w * 6 * days_91_plus


def solve_target_price(target_margin_pct: float, cost_fixed_rub: float, variable_rate_pct: float) -> Optional[float]:
    t = target_margin_pct / 100.0
    v = variable_rate_pct / 100.0
    denominator = 1 - v - t
    return None if denominator <= 0 else cost_fixed_rub / denominator


def calculate_row(
    row: pd.Series,
    reference_df: pd.DataFrame,
    scheme: str,
    fbs_weight_basis: str,
    fbs_logistics_profile: str,
    fbsm_logistics_profile: str,
    include_fbsm_return_logistics: bool,
    include_fbsm_defect_handling: bool,
    include_fbsm_excess_handling: bool,
    default_tax_system: str,
    manual_default_tax_pct: float,
) -> Dict:
    sku = str(row["Артикул"])
    name = str(row["Наименование товара"])
    cost = float(row["Себестоимость, ₽"])
    price = float(row["Цена продажи, ₽"])
    actual_weight = float(row["Вес факт, кг"])
    length = float(row["Длина, см"])
    width = float(row["Ширина, см"])
    height = float(row["Высота, см"])
    storage_days = float(row["Дней хранения FBSM"])
    ads_pct = float(row["Реклама, %"])
    row_tax_pct = float(row["Налог, %"])
    row_tax_system = str(row.get("Система налогообложения", "") or "")
    other_costs = float(row["Прочие расходы, ₽"])
    target_margin = float(row["Целевая маржа, %"])
    return_rate = float(row["Доля возвратов, %"])
    cancel_rate = float(row["Доля невыкупа/отмен, %"])

    tariff_group, lvl3_group, match_score, match_note = resolve_tariff_group(
        reference_df, name, row.get("Тарифная группа", ""), row.get("Товарная группа 3 уровня", "")
    )
    commission_pct = commission_rate_for_tariff_group(reference_df, tariff_group, scheme) or 0.0
    tax_pct, tax_system_label = resolve_tax_rate(row_tax_pct, row_tax_system, default_tax_system, manual_default_tax_pct)

    commission_rub = price * commission_pct / 100.0
    ad_cost_rub = price * ads_pct / 100.0
    tax_rub = price * tax_pct / 100.0

    logistics_to_buyer = reverse_logistics = storage_rub = defect_handling = excess_handling = billable_weight = 0.0
    if scheme == "FBS":
        billable_weight = calc_fbs_billable_weight(actual_weight, length, width, height, fbs_weight_basis)
        logistics_to_buyer = calc_fbs_delivery(billable_weight, fbs_logistics_profile)
    else:
        billable_weight = ceil_to(actual_weight, 0.1)
        logistics_to_buyer = calc_fbsm_delivery(billable_weight, fbsm_logistics_profile)
        if include_fbsm_return_logistics:
            reverse_logistics = logistics_to_buyer * (return_rate + cancel_rate) / 100.0
        storage_rub = calc_fbsm_storage(actual_weight, storage_days)
        defect_handling = 30.0 if include_fbsm_defect_handling else 0.0
        excess_handling = 30.0 if include_fbsm_excess_handling else 0.0

    mp_services_rub = logistics_to_buyer + reverse_logistics + storage_rub + defect_handling + excess_handling
    payout_before_seller_costs = price - commission_rub - mp_services_rub
    full_cost = cost + commission_rub + mp_services_rub + ad_cost_rub + tax_rub + other_costs
    profit = price - full_cost
    margin_on_revenue = (profit / price) if price else 0.0
    markup_on_full_cost = (price / full_cost - 1) if full_cost else 0.0

    target_price = target_margin_profit_rub = target_margin_pct_fact = None
    if target_margin > 0:
        fixed_costs = cost + logistics_to_buyer + reverse_logistics + storage_rub + defect_handling + excess_handling + other_costs
        variable_rate = commission_pct + ads_pct + tax_pct
        target_price = solve_target_price(target_margin, fixed_costs, variable_rate)
        if target_price:
            target_commission = target_price * commission_pct / 100.0
            target_ads = target_price * ads_pct / 100.0
            target_tax = target_price * tax_pct / 100.0
            target_full_cost = cost + mp_services_rub + other_costs + target_commission + target_ads + target_tax
            target_margin_profit_rub = target_price - target_full_cost
            target_margin_pct_fact = (target_margin_profit_rub / target_price) if target_price else None

    warnings = []
    if not tariff_group:
        warnings.append("Не определена тарифная группа")
    if commission_pct == 0:
        warnings.append("Комиссия не найдена")

    return {
        "Артикул": sku,
        "Наименование товара": name,
        "Схема": scheme,
        "Тарифная группа": tariff_group,
        "Товарная группа 3 уровня": lvl3_group,
        "Как определили категорию": match_note,
        "Совпадение, score": round(match_score or 0.0, 2) if match_score is not None else 0.0,
        "Налоговый режим": tax_system_label,
        "Цена продажи, ₽": round(price, 2),
        "Рекомендованная цена, ₽": round(target_price, 2) if target_price else None,
        "Себестоимость, ₽": round(cost, 2),
        "Комиссия, %": round(commission_pct / 100.0, 4),
        "Комиссия, ₽": round(commission_rub, 2),
        "Налог, %": round(tax_pct / 100.0, 4),
        "Налог, ₽": round(tax_rub, 2),
        "Вес для тарифа": round(billable_weight, 2),
        "Логистика до покупателя, ₽": round(logistics_to_buyer, 2),
        "Обратная логистика (ожидаемая), ₽": round(reverse_logistics, 2),
        "Хранение FBSM, ₽": round(storage_rub, 2),
        "Обработка брака, ₽": round(defect_handling, 2),
        "Обработка излишков, ₽": round(excess_handling, 2),
        "Реклама, ₽": round(ad_cost_rub, 2),
        "Прочие расходы, ₽": round(other_costs, 2),
        "Выплата до себестоимости, ₽": round(payout_before_seller_costs, 2),
        "Полная себестоимость, ₽": round(full_cost, 2),
        "Прибыль, ₽": round(profit, 2),
        "Маржа к выручке, %": round(margin_on_revenue, 4),
        "Наценка на полную себестоимость, %": round(markup_on_full_cost, 4),
        "Целевая маржа, %": round(target_margin / 100.0, 4),
        "Прибыль при рекомендованной цене, ₽": round(target_margin_profit_rub, 2) if target_margin_profit_rub is not None else None,
        "Маржа при рекомендованной цене, %": round(target_margin_pct_fact, 4) if target_margin_pct_fact is not None else None,
        "Комментарий": "; ".join(warnings),
    }


def build_export_workbook(result_df: pd.DataFrame) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Результаты"

    ws.append(list(result_df.columns))
    for row in result_df.itertuples(index=False):
        ws.append(list(row))

    header_fill = PatternFill("solid", fgColor="1F4E78")
    total_fill = PatternFill("solid", fgColor="E2F0D9")

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    money_cols = {c for c in result_df.columns if "₽" in c}
    pct_cols = {
        "Комиссия, %", "Налог, %", "Маржа к выручке, %", "Наценка на полную себестоимость, %",
        "Целевая маржа, %", "Маржа при рекомендованной цене, %",
    }

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            col_name = result_df.columns[cell.column - 1]
            if col_name in money_cols and isinstance(cell.value, (int, float)):
                cell.number_format = '#,##0.00_);[Red](#,##0.00)'
            elif col_name in pct_cols and isinstance(cell.value, (int, float)):
                cell.number_format = '0.0%'
            cell.alignment = Alignment(vertical="center", wrap_text=True)

    total_row = ws.max_row + 2
    ws.cell(total_row, 1).value = "ИТОГО / СРЕДНЕЕ"
    for idx, col_name in enumerate(result_df.columns, start=1):
        col_letter = get_column_letter(idx)
        cell = ws.cell(total_row, idx)
        cell.fill = total_fill
        cell.font = Font(bold=True)
        if col_name in money_cols:
            cell.value = f"=SUM({col_letter}2:{col_letter}{total_row - 2})"
            cell.number_format = '#,##0.00_);[Red](#,##0.00)'
        elif col_name in pct_cols:
            cell.value = f"=AVERAGE({col_letter}2:{col_letter}{total_row - 2})"
            cell.number_format = '0.0%'

    for col_idx, col_name in enumerate(result_df.columns, start=1):
        max_len = len(str(col_name))
        for row_idx in range(2, ws.max_row + 1):
            max_len = max(max_len, len(str(ws.cell(row_idx, col_idx).value or "")))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 12), 38)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    guide = wb.create_sheet("Пояснения")
    guide.append(["Показатель", "Описание"])
    for row in [
        ("Как определили категорию", "Показывает, было ли точное совпадение или автоподбор по названию товара."),
        ("Выплата до себестоимости, ₽", "Цена продажи минус комиссия и услуги маркетплейса."),
        ("Полная себестоимость, ₽", "Себестоимость товара плюс комиссия, логистика, хранение, реклама, налог и прочие расходы."),
        ("Маржа к выручке, %", "Прибыль / цена продажи."),
        ("Наценка на полную себестоимость, %", "Цена продажи / полная себестоимость - 1."),
        ("Рекомендованная цена, ₽", "Цена для достижения целевой маржи с учетом текущих параметров слева и параметров строки."),
    ]:
        guide.append(row)
    for cell in guide[1]:
        cell.fill = header_fill
        cell.font = Font(color="FFFFFF", bold=True)
    guide.column_dimensions["A"].width = 40
    guide.column_dimensions["B"].width = 110

    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)
    return stream.getvalue()


def render_overview_metrics(result_df: pd.DataFrame) -> None:
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("SKU", f"{len(result_df)}")
    c2.metric("Суммарная прибыль", f"{result_df['Прибыль, ₽'].sum():,.2f} ₽")
    c3.metric("Средняя маржа", f"{result_df['Маржа к выручке, %'].mean():.1%}")
    c4.metric(
        "Средняя рекомендованная цена",
        f"{result_df['Рекомендованная цена, ₽'].dropna().mean():,.2f} ₽" if result_df['Рекомендованная цена, ₽'].notna().any() else "—",
    )


def app() -> None:
    st.set_page_config(page_title="Спортмастер — юнит-экономика", page_icon="🏃", layout="wide")
    st.title("🏃 Спортмастер — юнит-экономика")
    st.caption("Массовая загрузка товаров через Excel • автоматическое определение категории • массовая Excel-выгрузка • FBS и FBSM")

    try:
        reference_df = read_reference()
    except Exception as e:
        st.error(str(e))
        st.stop()

    with st.sidebar:
        st.header("Параметры расчета")
        scheme = st.selectbox("Схема", ["FBS", "FBSM"], index=0)

        st.subheader("Налогообложение")
        default_tax_system = st.selectbox("Система налогообложения", list(TAX_SYSTEMS.keys()), index=0)
        manual_default_tax_pct = 0.0
        if default_tax_system == "Своя ставка":
            manual_default_tax_pct = st.number_input("Своя ставка налога, %", min_value=0.0, value=6.0, step=0.5)

        st.subheader("Логистика FBS")
        fbs_weight_basis_ui = st.selectbox(
            "Основа веса",
            ["Консервативно (макс из Почты/СДЭК)", "Почта России", "СДЭК", "Фактический"],
            index=0,
        )
        fbs_weight_basis_map = {
            "Консервативно (макс из Почты/СДЭК)": "Консервативно",
            "Почта России": "Почта России",
            "СДЭК": "СДЭК",
            "Фактический": "Фактический",
        }
        fbs_logistics_profile = st.selectbox("Тариф FBS", ["200 + 70", "220 + 90"], index=0)

        st.subheader("Параметры FBSM")
        fbsm_logistics_profile = st.selectbox("Тариф FBSM", ["35 / 75 / 35", "56 / 90 / 60"], index=0)
        include_fbsm_return_logistics = st.checkbox("Учитывать обратную логистику FBSM", value=True)
        include_fbsm_defect_handling = st.checkbox("Учитывать обработку брака 30 ₽/шт", value=False)
        include_fbsm_excess_handling = st.checkbox("Учитывать обработку излишков 30 ₽/шт", value=False)

    st.download_button(
        "Скачать шаблон товаров Excel",
        data=build_product_template_bytes(),
        file_name="Шаблон_товаров_Спортмастер.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=False,
    )
    products_file = st.file_uploader("Загрузите Excel-файл с товарами", type=["xlsx"])

    if products_file is None:
        st.info("Сначала скачайте шаблон, заполните товары и загрузите файл. Категория и тарифная группа определяются автоматически по названию товара.")
        return

    try:
        products_df = prepare_products(pd.read_excel(products_file))
    except Exception as e:
        st.error(f"Ошибка в файле товаров: {e}")
        st.stop()

    result_rows = [
        calculate_row(
            row,
            reference_df,
            scheme,
            fbs_weight_basis_map[fbs_weight_basis_ui],
            fbs_logistics_profile,
            fbsm_logistics_profile,
            include_fbsm_return_logistics,
            include_fbsm_defect_handling,
            include_fbsm_excess_handling,
            default_tax_system,
            manual_default_tax_pct,
        )
        for _, row in products_df.iterrows()
    ]
    result_df = pd.DataFrame(result_rows)
    display_df = result_df[[c for c in VISIBLE_COLUMNS if c in result_df.columns]].copy()
    render_overview_metrics(display_df)
    st.dataframe(display_df, use_container_width=True, hide_index=True, height=680)

    export_bytes = build_export_workbook(result_df)
    st.download_button(
        "Скачать результат в Excel",
        data=export_bytes,
        file_name=f"Спортмастер_юнит-экономика_{scheme}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )


if __name__ == "__main__":
    app()

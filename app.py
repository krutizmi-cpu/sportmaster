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
    "цена": "Цена продажи, ₽",
    "цена продажи": "Цена продажи, ₽",
    "price": "Цена продажи, ₽",
    "вес": "Вес факт, кг",
    "вес кг": "Вес факт, кг",
    "вес факт": "Вес факт, кг",
    "длина": "Длина, см",
    "ширина": "Ширина, см",
    "высота": "Высота, см",
    "дней хранения fbsm": "Дней хранения FBSM",
    "реклама": "Реклама, %",
    "система налогообложения": "Система налогообложения",
    "налог": "Налог, %",
    "прочие расходы": "Прочие расходы, ₽",
    "целевая маржа": "Целевая маржа, %",
    "доля возвратов": "Доля возвратов, %",
    "доля невыкупа отмен": "Доля невыкупа/отмен, %",
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
    "Вес для тарифа, кг",
    "Логистика до покупателя, ₽",
    "Обратная логистика (ожидаемая), ₽",
    "Хранение FBSM, ₽",
    "Обработка брака, ₽",
    "Обработка излишков, ₽",
    "Реклама, %",
    "Реклама, ₽",
    "Налоговый режим",
    "Ставка налога, %",
    "Налоговая база, ₽",
    "Налог, ₽",
    "Прочие расходы, ₽",
    "Выплата от МП, ₽",
    "Полная себестоимость, ₽",
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
    "велосипедный": "велосипед",
    "велосипедная": "велосипед",
    "велосипедные": "велосипед",
    "велозамок": "замок",
    "велошлем": "шлем",
    "велофонарь": "фонар",
    "велозвонок": "звонок",
    "велонасос": "насос",
    "вело": "велосипед",
    "кроссовки": "кроссовк",
    "ботинки": "ботинк",
    "перчатки": "перчатк",
    "варежки": "варежк",
    "мячи": "мяч",
    "гантели": "гантел",
    "горный": "горн",
    "беговые": "бег",
    "беговой": "бег",
    "лыжи": "лыж",
    "самокаты": "самокат",
    "замки": "замок",
    "фонари": "фонар",
    "держатели": "держател",
    "крылья": "крыл",
    "насосы": "насос",
    "звонки": "звонок",
    "шлемы": "шлем",
    "седла": "седло",
    "корзины": "корзин",
    "крепления": "креплен",
}

CATEGORY_OVERRIDE_RULES = [
    {"all": {"велосипед", "горн"}, "lvl3": "Велосипеды"},
    {"all": {"кроссовк", "бег"}, "lvl3": "Кроссовки для бега"},
    {"all": {"велосипед", "замок"}, "lvl3": "Замки для велосипеда"},
    {"all": {"замок"}, "any": {"велосипед", "вело"}, "lvl3": "Замки для велосипеда"},
    {"all": {"велосипед", "шлем"}, "lvl3": "Шлемы велосипедные"},
    {"all": {"шлем"}, "any": {"велосипед", "вело"}, "lvl3": "Шлемы велосипедные"},
    {"all": {"велосипед", "фонар"}, "lvl3": "Фонари для велосипеда"},
    {"all": {"фонар"}, "any": {"велосипед", "вело"}, "lvl3": "Фонари для велосипеда"},
    {"all": {"велосипед", "держател"}, "lvl3": "Держатели для велосипеда"},
    {"all": {"держател"}, "any": {"велосипед", "вело"}, "lvl3": "Держатели для велосипеда"},
    {"all": {"велосипед", "звонок"}, "lvl3": "Звонки для велосипеда"},
    {"all": {"звонок"}, "any": {"велосипед", "вело"}, "lvl3": "Звонки для велосипеда"},
    {"all": {"велосипед", "крыл"}, "lvl3": "Крылья для велосипеда"},
    {"all": {"крыл"}, "any": {"велосипед", "вело"}, "lvl3": "Крылья для велосипеда"},
    {"all": {"велосипед", "корзин"}, "lvl3": "Корзины для велосипеда"},
    {"all": {"велосипед", "насос"}, "lvl3": "Насосы для велосипеда"},
    {"all": {"насос"}, "any": {"велосипед", "вело"}, "lvl3": "Насосы для велосипеда"},
    {"all": {"велосипед", "седло"}, "lvl3": "Седла для велосипеда"},
    {"all": {"велосипед", "креплен"}, "lvl3": "Крепления для велосипеда"},
]


def normalize_text(value: str) -> str:
    s = str(value or "").strip().lower().replace("ё", "е")
    s = re.sub(r"[^0-9a-zа-я]+", " ", s)
    return " ".join(s.split())


def stem_token(token: str) -> str:
    t = normalize_text(token)
    if t in TOKEN_SYNONYMS:
        return TOKEN_SYNONYMS[t]
    for suffix in ["иями", "ями", "ами", "иях", "ия", "ья", "ье", "ий", "ый", "ой", "ая", "яя", "ое", "ее", "ые", "ие", "ам", "ям", "ах", "ях", "ом", "ем", "ую", "юю", "ов", "ев", "ей", "а", "я", "ы", "и", "е", "о", "у", "ю"]:
        if len(t) >= 6 and t.endswith(suffix):
            t = t[:-len(suffix)]
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
        raise FileNotFoundError(f"Не найден файл data/{DEFAULT_COMMISSIONS_FILENAME}. Положите его в repo в папку data.")

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

    combo_cols = ["Товарная группа 1 уровня", "Товарная группа 2 уровня", "Товарная группа 3 уровня", "Тарифная группа"]
    df["__search_text"] = df[combo_cols].fillna("").astype(str).agg(" ".join, axis=1)
    df["__search_norm"] = df["__search_text"].map(normalize_text)
    df["__tokens_combo"] = df["__search_text"].map(text_tokens)

    for col in ["Ставка комиссии FBSM, %", "Ставка комиссии FBS, %"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
    return df


def prepare_products(df: pd.DataFrame) -> pd.DataFrame:
    renamed = {c: COLUMN_ALIASES.get(normalize_text(c), str(c).strip()) for c in df.columns}
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
        df[c] = pd.to_numeric(df[c], errors="coerce")

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
        ["SM-001", "Кроссовки беговые мужские", 2500, 5990, 0.8, 32, 22, 12, "", 5, "ОСНО (22%)", "", 0, 20, 0, 0, "", ""],
        ["SM-002", "Велосипед горный", 18000, 32990, 14.5, 145, 25, 78, "", 7, "УСН доходы (6%)", "", 300, 18, 0, 0, "", ""],
        ["SM-003", "Велосипедный замок", 500, 1000, 0.3, 15, 10, 5, "", 0, "ОСНО (22%)", "", 0, 20, 0, 0, "", ""],
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
        ("Наименование товара", "Обязательное поле. Категория и тарифная группа определяются автоматически по названию."),
        ("Вес и габариты", "Обязательные поля: именно от них считается логистика. Цена на логистику не влияет."),
        ("Система налогообложения / Налог, %", "Можно оставить пустыми: приложение возьмет общий режим из левой панели."),
        ("Целевая маржа, %", "Если пусто или 0, берется значение по умолчанию из левой панели."),
        ("Тарифная группа / Товарная группа 3 уровня", "Обычно заполнять не нужно. Это ручное переопределение, если автоподбор не устроил."),
    ]
    for row in explanations:
        guide.append(row)
    for cell in guide[1]:
        cell.fill = header_fill
        cell.font = Font(color="FFFFFF", bold=True)
    guide.column_dimensions["A"].width = 42
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
    if pd.notna(row_tax_pct) and float(row_tax_pct) > 0:
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


def resolve_row_value(row: pd.Series, col_name: str, sidebar_default: float) -> float:
    value = row.get(col_name)
    if pd.isna(value) or value == "":
        return float(sidebar_default)
    try:
        value = float(value)
    except Exception:
        return float(sidebar_default)
    return float(sidebar_default if value == 0 else value)


def score_reference_match(name: str, row: pd.Series) -> float:
    name_norm = normalize_text(name)
    name_tokens = text_tokens(name)
    if not name_tokens:
        return 0.0

    best = 0.0
    for col in ["Товарная группа 3 уровня", "Тарифная группа", "Товарная группа 2 уровня"]:
        ref_norm = row[f"__norm_{col}"]
        ref_tokens = row[f"__tokens_{col}"]
        seq = SequenceMatcher(None, name_norm, ref_norm).ratio()
        inter = name_tokens & ref_tokens
        overlap = len(inter) / max(len(ref_tokens), 1)
        cover_name = len(inter) / max(len(name_tokens), 1)
        contains = 1.0 if ref_norm and (ref_norm in name_norm or any(tok in ref_norm for tok in name_tokens if len(tok) >= 5)) else 0.0
        score = max(seq * 0.35 + overlap * 0.40 + cover_name * 0.25, overlap * 0.75 + cover_name * 0.25, contains)
        best = max(best, score)

    combo_tokens = row["__tokens_combo"]
    combo_norm = row["__search_norm"]
    inter = name_tokens & combo_tokens
    combo_overlap = len(inter) / max(len(name_tokens), 1)
    combo_cover = len(inter) / max(len(combo_tokens), 1)
    prefix_hits = sum(1 for tok in name_tokens for ref_tok in combo_tokens if tok == ref_tok or tok.startswith(ref_tok) or ref_tok.startswith(tok))
    prefix_boost = min(prefix_hits * 0.05, 0.20)
    phrase_boost = 0.0

    if "велосипед" in name_tokens and "велоспорт" in combo_norm:
        phrase_boost += 0.15
    if {"велосипед", "замок"}.issubset(name_tokens) and "замки для велосипеда" in combo_norm:
        phrase_boost += 0.40
    if {"велосипед", "горн"}.issubset(name_tokens) and "велосипеды" in combo_norm:
        phrase_boost += 0.30
    if {"кроссовк", "бег"}.issubset(name_tokens) and "кроссовки для бега" in combo_norm:
        phrase_boost += 0.30
    if "аксессуары для велоспорта" in combo_norm and ("велосипед" in name_tokens or "вело" in name_norm):
        phrase_boost += 0.12

    best = max(best, combo_overlap * 0.72 + combo_cover * 0.10 + prefix_boost + phrase_boost)
    if inter and len(inter) >= 2:
        best = max(best, 0.80 + min(len(inter) * 0.02, 0.10))
    return min(best, 0.99)


def resolve_override_rule(reference_df: pd.DataFrame, product_name: str) -> Tuple[str, str, Optional[float], str] | None:
    name_tokens = text_tokens(product_name)
    if not name_tokens:
        return None
    for rule in CATEGORY_OVERRIDE_RULES:
        all_req = set(rule.get("all", set()))
        any_req = set(rule.get("any", set()))
        if all_req and not all_req.issubset(name_tokens):
            continue
        if any_req and not (any_req & name_tokens):
            continue
        rows = reference_df[reference_df["Товарная группа 3 уровня"] == rule["lvl3"]]
        if not rows.empty:
            row = rows.iloc[0]
            return str(row["Тарифная группа"]), str(row["Товарная группа 3 уровня"]), 0.99, "Правило по ключевым словам"
    return None


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

    override = resolve_override_rule(reference_df, product_name)
    if override is not None:
        return override

    for col, note in [
        ("Товарная группа 3 уровня", "Точное совпадение с ТГ3"),
        ("Тарифная группа", "Точное совпадение с тарифной группой"),
        ("Товарная группа 2 уровня", "Точное совпадение со 2 уровнем"),
    ]:
        rows = reference_df[reference_df[f"__norm_{col}"] == name_norm]
        if not rows.empty:
            row = rows.iloc[0]
            return str(row["Тарифная группа"]), str(row["Товарная группа 3 уровня"]), 1.0, note

    candidates = []
    for _, ref_row in reference_df.iterrows():
        score = score_reference_match(product_name, ref_row)
        if score > 0:
            candidates.append((score, str(ref_row["Тарифная группа"]), str(ref_row["Товарная группа 3 уровня"])))

    if not candidates:
        return "", "", None, "Не найдено"
    candidates.sort(key=lambda x: x[0], reverse=True)
    best_score, best_tariff, best_lvl3 = candidates[0]
    if best_score >= 0.42:
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


def tax_mode(label: str) -> str:
    label = str(label or "").lower()
    if "осно" in label:
        return "payout"
    if "доходы-расходы" in label:
        return "profit"
    if "доходы" in label:
        return "payout"
    return "revenue"


def calc_tax_amount(
    price: float,
    tax_pct: float,
    tax_system_label: str,
    payout_from_mp: float,
    profit_before_tax: float,
) -> Tuple[float, float]:
    rate = tax_pct / 100.0
    if rate <= 0 or price <= 0:
        return 0.0, 0.0

    mode = tax_mode(tax_system_label)
    if mode == "profit":
        base = max(profit_before_tax, 0.0)
    elif mode == "payout":
        base = max(payout_from_mp, 0.0)
    else:
        base = max(price, 0.0)
    return base, base * rate


def solve_target_price(
    target_margin_pct: float,
    cost_fixed_rub: float,
    commission_pct: float,
    ads_pct: float,
    tax_pct: float,
    tax_system_label: str,
) -> Optional[float]:
    t = target_margin_pct / 100.0
    c = commission_pct / 100.0
    a = ads_pct / 100.0
    r = tax_pct / 100.0
    mode = tax_mode(tax_system_label)

    if mode == "profit":
        # profit = (P*(1-c-a) - fixed)*(1-r)
        denominator = (1 - r) * (1 - c - a) - t
        return None if denominator <= 0 else (1 - r) * cost_fixed_rub / denominator

    if mode == "payout":
        # tax = r * (P*(1-c) - logistics/storage/handling)
        # profit = P - commission - ads - tax - fixed
        denominator = 1 - c - a - r * (1 - c) - t
        constant = cost_fixed_rub - r * cost_fixed_rub
        return None if denominator <= 0 else constant / denominator

    denominator = 1 - c - a - r - t
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
    sidebar_defaults: dict,
) -> Dict:
    sku = str(row["Артикул"])
    name = str(row["Наименование товара"])
    cost = float(row["Себестоимость, ₽"] or 0)
    price = float(row["Цена продажи, ₽"] or 0)
    actual_weight = float(row["Вес факт, кг"] or 0)
    length = float(row["Длина, см"] or 0)
    width = float(row["Ширина, см"] or 0)
    height = float(row["Высота, см"] or 0)
    storage_days = resolve_row_value(row, "Дней хранения FBSM", sidebar_defaults["Дней хранения FBSM"])
    ads_pct = resolve_row_value(row, "Реклама, %", sidebar_defaults["Реклама, %"])
    row_tax_pct = row["Налог, %"]
    row_tax_system = str(row.get("Система налогообложения", "") or "")
    other_costs = resolve_row_value(row, "Прочие расходы, ₽", sidebar_defaults["Прочие расходы, ₽"])
    target_margin = resolve_row_value(row, "Целевая маржа, %", sidebar_defaults["Целевая маржа, %"])
    return_rate = resolve_row_value(row, "Доля возвратов, %", sidebar_defaults["Доля возвратов, %"])
    cancel_rate = resolve_row_value(row, "Доля невыкупа/отмен, %", sidebar_defaults["Доля невыкупа/отмен, %"])

    tariff_group, lvl3_group, match_score, match_note = resolve_tariff_group(
        reference_df, name, row.get("Тарифная группа", ""), row.get("Товарная группа 3 уровня", "")
    )
    commission_pct = commission_rate_for_tariff_group(reference_df, tariff_group, scheme) or 0.0
    tax_pct, tax_system_label = resolve_tax_rate(row_tax_pct, row_tax_system, default_tax_system, manual_default_tax_pct)

    commission_rub = price * commission_pct / 100.0
    ad_cost_rub = price * ads_pct / 100.0
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
    payout_from_mp = price - commission_rub - mp_services_rub
    profit_before_tax = payout_from_mp - cost - ad_cost_rub - other_costs
    tax_base_rub, tax_rub = calc_tax_amount(
        price=price,
        tax_pct=tax_pct,
        tax_system_label=tax_system_label,
        payout_from_mp=payout_from_mp,
        profit_before_tax=profit_before_tax,
    )

    full_cost = cost + commission_rub + mp_services_rub + ad_cost_rub + tax_rub + other_costs
    profit = price - full_cost
    margin_on_revenue_pct = (profit / price * 100.0) if price else 0.0
    markup_on_full_cost_pct = (price / full_cost - 1.0) * 100.0 if full_cost else 0.0

    target_price = target_profit_rub = target_margin_pct_fact = None
    fixed_costs = cost + logistics_to_buyer + reverse_logistics + storage_rub + defect_handling + excess_handling + other_costs
    if target_margin > 0:
        target_price = solve_target_price(
            target_margin_pct=target_margin,
            cost_fixed_rub=fixed_costs,
            commission_pct=commission_pct,
            ads_pct=ads_pct,
            tax_pct=tax_pct,
            tax_system_label=tax_system_label,
        )
        if target_price is not None:
            target_commission = target_price * commission_pct / 100.0
            target_ads = target_price * ads_pct / 100.0
            target_payout = target_price - target_commission - mp_services_rub
            target_profit_before_tax = target_payout - cost - target_ads - other_costs
            _, target_tax = calc_tax_amount(
                price=target_price,
                tax_pct=tax_pct,
                tax_system_label=tax_system_label,
                payout_from_mp=target_payout,
                profit_before_tax=target_profit_before_tax,
            )
            target_full_cost = cost + mp_services_rub + other_costs + target_commission + target_ads + target_tax
            target_profit_rub = target_price - target_full_cost
            target_margin_pct_fact = (target_profit_rub / target_price * 100.0) if target_price else None

    warnings = []
    if not tariff_group:
        warnings.append("Не определена тарифная группа")
    if commission_pct == 0:
        warnings.append("Комиссия не найдена")
    if target_margin <= 0:
        warnings.append("Целевая маржа не задана — рекомендованная цена не рассчитана")
    elif target_price is None:
        warnings.append("Рекомендованная цена не считается: слишком высокая сумма комиссии, рекламы, налога и целевой маржи")
    if price <= 0:
        warnings.append("Цена продажи 0")
    if actual_weight <= 0:
        warnings.append("Вес товара 0")

    return {
        "Артикул": sku,
        "Наименование товара": name,
        "Схема": scheme,
        "Тарифная группа": tariff_group,
        "Товарная группа 3 уровня": lvl3_group,
        "Как определили категорию": match_note,
        "Налоговый режим": tax_system_label,
        "Цена продажи, ₽": round(price, 2),
        "Рекомендованная цена, ₽": round(target_price, 2) if target_price is not None else None,
        "Себестоимость, ₽": round(cost, 2),
        "Комиссия, %": round(commission_pct, 2),
        "Комиссия, ₽": round(commission_rub, 2),
        "Вес для тарифа, кг": round(billable_weight, 2),
        "Логистика до покупателя, ₽": round(logistics_to_buyer, 2),
        "Обратная логистика (ожидаемая), ₽": round(reverse_logistics, 2),
        "Хранение FBSM, ₽": round(storage_rub, 2),
        "Обработка брака, ₽": round(defect_handling, 2),
        "Обработка излишков, ₽": round(excess_handling, 2),
        "Реклама, %": round(ads_pct, 2),
        "Реклама, ₽": round(ad_cost_rub, 2),
        "Ставка налога, %": round(tax_pct, 2),
        "Налоговая база, ₽": round(tax_base_rub, 2),
        "Налог, ₽": round(tax_rub, 2),
        "Прочие расходы, ₽": round(other_costs, 2),
        "Выплата от МП, ₽": round(payout_from_mp, 2),
        "Полная себестоимость, ₽": round(full_cost, 2),
        "Прибыль, ₽": round(profit, 2),
        "Маржа к выручке, %": round(margin_on_revenue_pct, 2),
        "Наценка на полную себестоимость, %": round(markup_on_full_cost_pct, 2),
        "Целевая маржа, %": round(target_margin, 2),
        "Прибыль при рекомендованной цене, ₽": round(target_profit_rub, 2) if target_profit_rub is not None else None,
        "Маржа при рекомендованной цене, %": round(target_margin_pct_fact, 2) if target_margin_pct_fact is not None else None,
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
        "Комиссия, %", "Реклама, %", "Ставка налога, %", "Маржа к выручке, %",
        "Наценка на полную себестоимость, %", "Целевая маржа, %", "Маржа при рекомендованной цене, %",
    }

    for row in ws.iter_rows(min_row=2):
        for cell in row:
            col_name = result_df.columns[cell.column - 1]
            if col_name in money_cols and isinstance(cell.value, (int, float)):
                cell.number_format = '#,##0.00_);[Red](#,##0.00)'
            elif col_name in pct_cols and isinstance(cell.value, (int, float)):
                cell.number_format = '0.00'
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
            cell.number_format = '0.00'

    for col_idx, col_name in enumerate(result_df.columns, start=1):
        max_len = len(str(col_name))
        for row_idx in range(2, ws.max_row + 1):
            max_len = max(max_len, len(str(ws.cell(row_idx, col_idx).value or "")))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 12), 40)

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    guide = wb.create_sheet("Пояснения")
    guide.append(["Показатель", "Описание"])
    for row in [
        ("Выплата от МП, ₽", "Что остается после комиссии и услуг маркетплейса, но до себестоимости, рекламы, налога и прочих расходов."),
        ("Налоговая база, ₽", "База, с которой рассчитан налог по выбранному режиму."),
        ("Полная себестоимость, ₽", "Себестоимость товара плюс комиссия, логистика, хранение, реклама, налог и прочие расходы."),
        ("Маржа к выручке, %", "Прибыль / цена продажи * 100."),
        ("Наценка на полную себестоимость, %", "(Цена продажи / полная себестоимость - 1) * 100."),
    ]:
        guide.append(row)
    for cell in guide[1]:
        cell.fill = header_fill
        cell.font = Font(color="FFFFFF", bold=True)
    guide.column_dimensions["A"].width = 40
    guide.column_dimensions["B"].width = 115

    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)
    return stream.getvalue()


def render_overview_metrics(result_df: pd.DataFrame) -> None:
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("SKU", f"{len(result_df)}")
    c2.metric("Суммарная прибыль", f"{result_df['Прибыль, ₽'].sum():,.2f} ₽")
    c3.metric("Средняя маржа", f"{result_df['Маржа к выручке, %'].mean():.1f}%")
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

        st.subheader("Параметры по умолчанию")
        default_target_margin_pct = st.number_input("Целевая маржа по умолчанию, %", min_value=0.0, value=20.0, step=1.0)
        default_ads_pct = st.number_input("Реклама по умолчанию, %", min_value=0.0, value=0.0, step=0.5)
        default_storage_days = st.number_input("Дней хранения FBSM по умолчанию", min_value=0.0, value=0.0, step=1.0)
        default_other_costs = st.number_input("Прочие расходы по умолчанию, ₽", min_value=0.0, value=0.0, step=50.0)
        default_returns_pct = st.number_input("Доля возвратов по умолчанию, %", min_value=0.0, value=0.0, step=0.5)
        default_cancel_pct = st.number_input("Доля невыкупа/отмен по умолчанию, %", min_value=0.0, value=0.0, step=0.5)

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

    sidebar_defaults = {
        "Дней хранения FBSM": default_storage_days,
        "Реклама, %": default_ads_pct,
        "Прочие расходы, ₽": default_other_costs,
        "Целевая маржа, %": default_target_margin_pct,
        "Доля возвратов, %": default_returns_pct,
        "Доля невыкупа/отмен, %": default_cancel_pct,
    }

    st.download_button(
        "Скачать шаблон товаров Excel",
        data=build_product_template_bytes(),
        file_name="Шаблон_товаров_Спортмастер.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    products_file = st.file_uploader("Загрузите Excel-файл с товарами", type=["xlsx"])
    st.caption("Логистика пересчитывается от веса и габаритов. Изменение цены влияет на комиссию, рекламу, налог, прибыль и рекомендованную цену, но не на сам тариф логистики.")

    if products_file is None:
        st.info("Сначала скачайте шаблон, заполните товары и загрузите файл.")
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
            sidebar_defaults,
        )
        for _, row in products_df.iterrows()
    ]
    result_df = pd.DataFrame(result_rows)
    display_df = result_df[[c for c in VISIBLE_COLUMNS if c in result_df.columns]].copy()
    render_overview_metrics(display_df)
    st.dataframe(display_df, use_container_width=True, hide_index=True, height=700)

    export_bytes = build_export_workbook(display_df)
    st.download_button(
        "Скачать результат в Excel",
        data=export_bytes,
        file_name=f"Спортмастер_юнит-экономика_{scheme}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
    )


if __name__ == "__main__":
    app()

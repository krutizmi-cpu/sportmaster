from __future__ import annotations

import io
import math
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
DEFAULT_COMMISSIONS_FILE = DATA_DIR / "sportmaster_commissions_2026-02-01.xlsx"
PRODUCT_TEMPLATE_FILE = TEMPLATES_DIR / "Шаблон_товаров_Спортмастер.xlsx"

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
    "Тарифная группа",
    "Товарная группа 3 уровня",
    "Дней хранения FBSM",
    "Реклама, %",
    "Налог, %",
    "Прочие расходы, ₽",
    "Целевая маржа, %",
    "Доля возвратов, %",
    "Доля невыкупа/отмен, %",
]

COLUMN_ALIASES = {
    "sku": "Артикул",
    "артикул": "Артикул",
    "наименование": "Наименование товара",
    "наименование товара": "Наименование товара",
    "товар": "Наименование товара",
    "себестоимость": "Себестоимость, ₽",
    "себестоимость, руб": "Себестоимость, ₽",
    "цена": "Цена продажи, ₽",
    "цена продажи": "Цена продажи, ₽",
    "price": "Цена продажи, ₽",
    "вес": "Вес факт, кг",
    "вес факт": "Вес факт, кг",
    "вес факт, кг": "Вес факт, кг",
    "длина": "Длина, см",
    "ширина": "Ширина, см",
    "высота": "Высота, см",
    "тарифная группа": "Тарифная группа",
    "товарная группа 3 уровня": "Товарная группа 3 уровня",
    "дней хранения fbsm": "Дней хранения FBSM",
    "реклама": "Реклама, %",
    "налог": "Налог, %",
    "прочие расходы": "Прочие расходы, ₽",
    "целевая маржа": "Целевая маржа, %",
    "доля возвратов": "Доля возвратов, %",
    "доля невыкупа/отмен": "Доля невыкупа/отмен, %",
}


def normalize_text(value: str) -> str:
    s = str(value or "").strip().lower()
    for ch in ["ё", "/", "-", ",", ".", "(", ")", '"', "'", "+"]:
        s = s.replace(ch, "е" if ch == "ё" else " ")
    return " ".join(s.split())


def ceil_to(value: float, step: float) -> float:
    if value <= 0:
        return 0.0
    return math.ceil(value / step) * step


def read_reference(uploaded_file=None) -> pd.DataFrame:
    src = uploaded_file if uploaded_file is not None else DEFAULT_COMMISSIONS_FILE
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
    for col in ["Товарная группа 1 уровня", "Товарная группа 2 уровня", "Товарная группа 3 уровня", "Тарифная группа"]:
        df[col] = df[col].astype(str).str.strip()
        df[f"__norm_{col}"] = df[col].map(normalize_text)
    for col in ["Ставка комиссии FBSM, %", "Ставка комиссии FBS, %"]:
        df[col] = pd.to_numeric(df[col], errors="coerce").fillna(0.0)
    return df


def prepare_products(df: pd.DataFrame) -> pd.DataFrame:
    renamed = {}
    for c in df.columns:
        key = normalize_text(c)
        renamed[c] = COLUMN_ALIASES.get(key, str(c).strip())
    df = df.rename(columns=renamed).copy()
    missing = [c for c in PRODUCT_REQUIRED_COLS if c not in df.columns]
    if missing:
        raise ValueError(f"В файле товаров не хватает колонок: {', '.join(missing)}")
    for c in PRODUCT_OPTIONAL_COLS:
        if c not in df.columns:
            df[c] = None
    numeric_cols = [
        "Себестоимость, ₽",
        "Цена продажи, ₽",
        "Вес факт, кг",
        "Длина, см",
        "Ширина, см",
        "Высота, см",
        "Дней хранения FBSM",
        "Реклама, %",
        "Налог, %",
        "Прочие расходы, ₽",
        "Целевая маржа, %",
        "Доля возвратов, %",
        "Доля невыкупа/отмен, %",
    ]
    for c in numeric_cols:
        df[c] = pd.to_numeric(df[c], errors="coerce").fillna(0.0)
    return df


def commission_rate_for_tariff_group(reference_df: pd.DataFrame, tariff_group: str, scheme: str) -> Optional[float]:
    if not tariff_group:
        return None
    col = "Ставка комиссии FBSM, %" if scheme == "FBSM" else "Ставка комиссии FBS, %"
    rows = reference_df[reference_df["Тарифная группа"] == tariff_group]
    if rows.empty:
        return None
    rate = float(rows.iloc[0][col])
    if rate <= 1:
        rate *= 100
    return rate


def resolve_tariff_group(reference_df: pd.DataFrame, product_name: str, manual_tariff_group: str, manual_group3: str) -> Tuple[str, str, Optional[float], str]:
    name_norm = normalize_text(product_name)
    manual_tariff_group = str(manual_tariff_group or "").strip()
    manual_group3 = str(manual_group3 or "").strip()

    if manual_tariff_group:
        rate_row = reference_df[reference_df["Тарифная группа"] == manual_tariff_group]
        if not rate_row.empty:
            group3 = str(rate_row.iloc[0]["Товарная группа 3 уровня"])
            return manual_tariff_group, group3, 1.0, "Тарифная группа из файла товаров"

    if manual_group3:
        rows = reference_df[reference_df["Товарная группа 3 уровня"] == manual_group3]
        if not rows.empty:
            tariff_group = str(rows.iloc[0]["Тарифная группа"])
            return tariff_group, manual_group3, 1.0, "Товарная группа 3 уровня из файла товаров"

    exact_lvl3 = reference_df[reference_df["__norm_Товарная группа 3 уровня"] == name_norm]
    if not exact_lvl3.empty:
        row = exact_lvl3.iloc[0]
        return str(row["Тарифная группа"]), str(row["Товарная группа 3 уровня"]), 1.0, "Точное совпадение с ТГ3"

    exact_tariff = reference_df[reference_df["__norm_Тарифная группа"] == name_norm]
    if not exact_tariff.empty:
        row = exact_tariff.iloc[0]
        return str(row["Тарифная группа"]), str(row["Товарная группа 3 уровня"]), 1.0, "Точное совпадение с тарифной группой"

    candidates = []
    for _, row in reference_df.iterrows():
        lvl3_norm = row["__norm_Товарная группа 3 уровня"]
        tariff_norm = row["__norm_Тарифная группа"]
        score_lvl3 = SequenceMatcher(None, name_norm, lvl3_norm).ratio()
        score_tariff = SequenceMatcher(None, name_norm, tariff_norm).ratio()
        contains_score = 0.0
        if lvl3_norm and (lvl3_norm in name_norm or name_norm in lvl3_norm):
            contains_score = max(contains_score, 0.92)
        if tariff_norm and (tariff_norm in name_norm or name_norm in tariff_norm):
            contains_score = max(contains_score, 0.90)
        score = max(score_lvl3, score_tariff, contains_score)
        candidates.append((score, str(row["Тарифная группа"]), str(row["Товарная группа 3 уровня"])))
    candidates.sort(reverse=True)
    best_score, best_tariff, best_lvl3 = candidates[0]
    if best_score >= 0.78:
        return best_tariff, best_lvl3, best_score, f"Автоподбор ({best_score:.2f})"
    return "", "", best_score, "Не найдено"


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
    if profile == "220 + 90":
        base, extra = 220.0, 90.0
    else:
        base, extra = 200.0, 70.0
    if weight_kg <= 2:
        return base
    return base + (weight_kg - 2) * extra


def calc_fbsm_delivery(weight_kg: float, profile: str) -> float:
    w = ceil_to(weight_kg, 0.1)
    if w <= 0:
        return 0.0
    if profile == "56 / 90 / 60":
        up05, up10, extra = 56.0, 90.0, 60.0
    else:
        up05, up10, extra = 35.0, 75.0, 35.0
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
    if denominator <= 0:
        return None
    return cost_fixed_rub / denominator


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
    tax_pct = float(row["Налог, %"])
    other_costs = float(row["Прочие расходы, ₽"])
    target_margin = float(row["Целевая маржа, %"])
    return_rate = float(row["Доля возвратов, %"])
    cancel_rate = float(row["Доля невыкупа/отмен, %"])

    tariff_group, lvl3_group, match_score, match_note = resolve_tariff_group(
        reference_df, name, row.get("Тарифная группа", ""), row.get("Товарная группа 3 уровня", "")
    )
    commission_pct = commission_rate_for_tariff_group(reference_df, tariff_group, scheme)
    if commission_pct is None:
        commission_pct = 0.0

    commission_rub = price * commission_pct / 100.0
    ad_cost_rub = price * ads_pct / 100.0
    tax_rub = price * tax_pct / 100.0

    logistics_to_buyer = 0.0
    reverse_logistics = 0.0
    storage_rub = 0.0
    defect_handling = 0.0
    excess_handling = 0.0
    billable_weight = 0.0
    logistics_note = ""

    if scheme == "FBS":
        billable_weight = calc_fbs_billable_weight(actual_weight, length, width, height, fbs_weight_basis)
        logistics_to_buyer = calc_fbs_delivery(billable_weight, fbs_logistics_profile)
        logistics_note = f"FBS: {fbs_logistics_profile}, вес для тарифа = {billable_weight:.0f} кг"
        reverse_logistics = 0.0  # по текущему описанию обратная логистика не тарифицируется
    else:
        billable_weight = ceil_to(actual_weight, 0.1)
        logistics_to_buyer = calc_fbsm_delivery(billable_weight, fbsm_logistics_profile)
        if include_fbsm_return_logistics:
            reverse_logistics = logistics_to_buyer * (return_rate + cancel_rate) / 100.0
        storage_rub = calc_fbsm_storage(actual_weight, storage_days)
        defect_handling = 30.0 if include_fbsm_defect_handling else 0.0
        excess_handling = 30.0 if include_fbsm_excess_handling else 0.0
        logistics_note = f"FBSM: {fbsm_logistics_profile}, вес для тарифа = {billable_weight:.1f} кг"

    mp_services_rub = logistics_to_buyer + reverse_logistics + storage_rub + defect_handling + excess_handling
    payout_before_seller_costs = price - commission_rub - mp_services_rub
    full_cost = cost + commission_rub + mp_services_rub + ad_cost_rub + tax_rub + other_costs
    profit = price - full_cost
    margin_on_revenue = (profit / price) if price else 0.0
    markup_on_cost = (profit / full_cost) if full_cost else 0.0

    target_price = None
    if target_margin > 0:
        fixed_costs = cost + logistics_to_buyer + reverse_logistics + storage_rub + defect_handling + excess_handling + other_costs
        variable_rate = commission_pct + ads_pct + tax_pct
        target_price = solve_target_price(target_margin, fixed_costs, variable_rate)

    warnings = []
    if not tariff_group:
        warnings.append("Не определена тарифная группа")
    if commission_pct == 0:
        warnings.append("Комиссия не найдена")
    if scheme == "FBSM" and storage_days > 60 and storage_rub == 0:
        warnings.append("Проверьте хранение")

    return {
        "Артикул": sku,
        "Наименование товара": name,
        "Схема": scheme,
        "Тарифная группа": tariff_group,
        "Товарная группа 3 уровня": lvl3_group,
        "Как определили категорию": match_note,
        "Совпадение, score": round(match_score or 0.0, 2) if match_score is not None else 0.0,
        "Цена продажи, ₽": round(price, 2),
        "Себестоимость, ₽": round(cost, 2),
        "Комиссия, %": round(commission_pct / 100.0, 4),
        "Комиссия, ₽": round(commission_rub, 2),
        "Вес для тарифа": round(billable_weight, 2),
        "Логистика до покупателя, ₽": round(logistics_to_buyer, 2),
        "Обратная логистика (ожидаемая), ₽": round(reverse_logistics, 2),
        "Хранение FBSM, ₽": round(storage_rub, 2),
        "Обработка брака, ₽": round(defect_handling, 2),
        "Обработка излишков, ₽": round(excess_handling, 2),
        "Реклама, ₽": round(ad_cost_rub, 2),
        "Налог, ₽": round(tax_rub, 2),
        "Прочие расходы, ₽": round(other_costs, 2),
        "Выплата до себестоимости, ₽": round(payout_before_seller_costs, 2),
        "Полная себестоимость, ₽": round(full_cost, 2),
        "Прибыль, ₽": round(profit, 2),
        "Маржа к выручке, %": round(margin_on_revenue, 4),
        "Наценка на затраты, %": round(markup_on_cost, 4),
        "Целевая маржа, %": round(target_margin / 100.0, 4),
        "Рекомендованная цена, ₽": round(target_price, 2) if target_price else None,
        "Логистическая модель": logistics_note,
        "Комментарий": "; ".join(warnings),
    }


def build_export_workbook(result_df: pd.DataFrame) -> bytes:
    wb = Workbook()
    ws = wb.active
    ws.title = "Результаты"

    headers = list(result_df.columns)
    ws.append(headers)
    for row in result_df.itertuples(index=False):
        ws.append(list(row))

    header_fill = PatternFill("solid", fgColor="1F4E78")
    summary_fill = PatternFill("solid", fgColor="E2F0D9")
    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = Font(color="FFFFFF", bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    money_cols = {c for c in headers if "₽" in c}
    pct_cols = {"Комиссия, %", "Маржа к выручке, %", "Наценка на затраты, %", "Целевая маржа, %"}
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            col_name = headers[cell.column - 1]
            if col_name in money_cols and isinstance(cell.value, (int, float)):
                cell.number_format = '#,##0.00_);[Red](#,##0.00)'
            elif col_name in pct_cols and isinstance(cell.value, (int, float)):
                cell.number_format = '0.0%'
            cell.alignment = Alignment(vertical="center", wrap_text=True)

    for col_idx, col_name in enumerate(headers, start=1):
        max_len = max(len(str(col_name)), *(len(str(ws.cell(r, col_idx).value or "")) for r in range(2, ws.max_row + 1)))
        ws.column_dimensions[get_column_letter(col_idx)].width = min(max(max_len + 2, 14), 34)

    total_row = ws.max_row + 2
    ws.cell(total_row, 1).value = "ИТОГО / СРЕДНЕЕ"
    for cell in ws[total_row]:
        cell.fill = summary_fill
        cell.font = Font(bold=True)
    for idx, col_name in enumerate(headers, start=1):
        col_letter = get_column_letter(idx)
        if col_name in money_cols:
            ws.cell(total_row, idx).value = f"=SUM({col_letter}2:{col_letter}{total_row-2})"
            ws.cell(total_row, idx).number_format = '#,##0.00_);[Red](#,##0.00)'
        elif col_name in pct_cols:
            ws.cell(total_row, idx).value = f"=AVERAGE({col_letter}2:{col_letter}{total_row-2})"
            ws.cell(total_row, idx).number_format = '0.0%'

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    guide = wb.create_sheet("Как читать")
    guide.append(["Показатель", "Пояснение"])
    guide.append(["Выплата до себестоимости, ₽", "Цена продажи минус комиссия Спортмастера и услуги маркетплейса."])
    guide.append(["Полная себестоимость, ₽", "Себестоимость товара + комиссия + услуги МП + реклама + налог + прочие расходы."])
    guide.append(["Маржа к выручке, %", "Прибыль / цена продажи."])
    guide.append(["Наценка на затраты, %", "Прибыль / полная себестоимость."])
    guide.append(["Рекомендованная цена, ₽", "Цена для достижения целевой маржи с учетом текущей структуры затрат."])
    for c in guide[1]:
        c.fill = header_fill
        c.font = Font(color="FFFFFF", bold=True)
    guide.column_dimensions['A'].width = 30
    guide.column_dimensions['B'].width = 90

    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)
    return stream.getvalue()


def render_single_result(result: Dict) -> None:
    st.success("Расчет готов")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Выплата до себестоимости", f"{result['Выплата до себестоимости, ₽']:.2f} ₽")
    c2.metric("Прибыль", f"{result['Прибыль, ₽']:.2f} ₽")
    c3.metric("Маржа к выручке", f"{result['Маржа к выручке, %']:.1%}")
    c4.metric("Полная себестоимость", f"{result['Полная себестоимость, ₽']:.2f} ₽")

    detail_cols = [
        "Комиссия, ₽", "Логистика до покупателя, ₽", "Обратная логистика (ожидаемая), ₽",
        "Хранение FBSM, ₽", "Обработка брака, ₽", "Обработка излишков, ₽",
        "Реклама, ₽", "Налог, ₽", "Прочие расходы, ₽"
    ]
    df = pd.DataFrame({
        "Статья": detail_cols,
        "Сумма, ₽": [result[c] for c in detail_cols]
    })
    st.subheader("Структура расходов")
    st.dataframe(df, use_container_width=True, hide_index=True)
    st.bar_chart(df.set_index("Статья")["Сумма, ₽"])

    with st.expander("Детали"):
        st.write(pd.DataFrame([result]).T)


def app() -> None:
    st.set_page_config(page_title="Спортмастер — юнит-экономика", page_icon="🏃", layout="wide")
    st.title("🏃 Спортмастер — юнит-экономика")
    st.caption("FBS и FBSM • комиссии 01.02.2026 • автоматический расчет логистики и Excel-выгрузка")

    with st.sidebar:
        st.header("Настройки расчета")
        page = st.radio("Раздел", ["Калькулятор", "Пакетный расчет", "Справка"], index=0)
        scheme = st.selectbox("Схема", ["FBS", "FBSM"], index=0)
        ref_upload = st.file_uploader("Файл комиссий Sportmaster (опционально)", type=["xlsx"])

        st.markdown("---")
        st.subheader("Логистика FBS")
        fbs_weight_basis = st.selectbox("Основа веса", ["Консервативно (макс из Почты/СДЭК)", "Почта России", "СДЭК", "Фактический"], index=0)
        fbs_weight_basis_map = {
            "Консервативно (макс из Почты/СДЭК)": "Консервативно",
            "Почта России": "Почта России",
            "СДЭК": "СДЭК",
            "Фактический": "Фактический",
        }
        fbs_logistics_profile = st.selectbox("Тариф FBS", ["200 + 70", "220 + 90"], index=0)

        st.subheader("Логистика FBSM")
        fbsm_logistics_profile = st.selectbox("Тариф FBSM", ["35 / 75 / 35", "56 / 90 / 60"], index=0)
        include_fbsm_return_logistics = st.checkbox("Учитывать ожидаемую обратную логистику FBSM", value=True)
        include_fbsm_defect_handling = st.checkbox("Добавлять обработку брака 30 ₽/шт", value=False)
        include_fbsm_excess_handling = st.checkbox("Добавлять обработку излишков 30 ₽/шт", value=False)

        st.markdown("---")
        if PRODUCT_TEMPLATE_FILE.exists():
            with open(PRODUCT_TEMPLATE_FILE, "rb") as f:
                st.download_button("Скачать шаблон товаров", f.read(), file_name=PRODUCT_TEMPLATE_FILE.name)
        if DEFAULT_COMMISSIONS_FILE.exists():
            with open(DEFAULT_COMMISSIONS_FILE, "rb") as f:
                st.download_button("Скачать файл комиссий в repo", f.read(), file_name=DEFAULT_COMMISSIONS_FILE.name)

    try:
        reference_df = read_reference(ref_upload)
    except Exception as e:
        st.error(f"Не удалось прочитать файл комиссий: {e}")
        st.stop()

    if page == "Справка":
        st.subheader("Что исправлено относительно старого app.py")
        st.markdown(
            """
            - комиссия больше **не вводится руками**, а подтягивается из официального файла Sportmaster;
            - добавлены **FBS и FBSM**;
            - логистика считается по правилам Sportmaster, а не вводится вручную;
            - хранение FBSM считается по периодам **0–60 / 61–90 / 91+**;
            - учтены **обработка брака** и **обработка излишков** для FBSM;
            - вынесены отдельные поля под **налог, рекламу, прочие расходы**;
            - добавлены **Excel-шаблон и Excel-выгрузка**.
            """
        )
        st.subheader("Что нашел по скрытым комиссиям")
        st.markdown(
            """
            По открытым страницам Sportmaster я нашел:
            - агентскую комиссию по категориям;
            - услуги доставки FBS/FBSM;
            - хранение FBSM;
            - обработку брака и излишков FBSM;
            - рекламные услуги и дополнительные услуги как отдельные документы в отчетности.

            Отдельной подтвержденной услуги типа **«быстрый вывод»** в доступных страницах не нашел, поэтому ее **не зашивал как обязательную**.
            Если у вас в договоре/УПД она есть, ее лучше добавлять через колонку **«Прочие расходы, ₽»** или отдельное поле в следующей версии.
            """
        )
        st.subheader("Справочник тарифных групп")
        tariff_table = reference_df[["Тарифная группа", "Ставка комиссии FBSM, %", "Ставка комиссии FBS, %"]].drop_duplicates().sort_values("Тарифная группа")
        tariff_table["Ставка комиссии FBSM, %"] = tariff_table["Ставка комиссии FBSM, %"].apply(lambda x: x * 100 if x <= 1 else x)
        tariff_table["Ставка комиссии FBS, %"] = tariff_table["Ставка комиссии FBS, %"].apply(lambda x: x * 100 if x <= 1 else x)
        st.dataframe(tariff_table, use_container_width=True, height=500)
        return

    if page == "Калькулятор":
        st.header("Калькулятор 1 SKU")
        c1, c2, c3 = st.columns(3)
        with c1:
            sku = st.text_input("Артикул", value="SKU-001")
            name = st.text_input("Название товара", value="Кроссовки беговые мужские")
            cost = st.number_input("Себестоимость, ₽", min_value=0.0, value=2500.0, step=100.0)
            price = st.number_input("Цена продажи, ₽", min_value=0.0, value=5990.0, step=100.0)
        with c2:
            weight = st.number_input("Вес факт, кг", min_value=0.0, value=0.8, step=0.1)
            length = st.number_input("Длина, см", min_value=0.0, value=32.0, step=1.0)
            width = st.number_input("Ширина, см", min_value=0.0, value=22.0, step=1.0)
            height = st.number_input("Высота, см", min_value=0.0, value=12.0, step=1.0)
        with c3:
            manual_tariff = st.text_input("Тарифная группа (если знаете)", value="Обувь для взрослых и детей")
            manual_lvl3 = st.text_input("Товарная группа 3 уровня (если знаете)", value="")
            storage_days = st.number_input("Дней хранения FBSM", min_value=0.0, value=0.0, step=1.0)
            ads_pct = st.number_input("Реклама, %", min_value=0.0, value=5.0, step=0.5)
            tax_pct = st.number_input("Налог, %", min_value=0.0, value=6.0, step=0.5)
            other_costs = st.number_input("Прочие расходы, ₽", min_value=0.0, value=0.0, step=50.0)
            target_margin = st.number_input("Целевая маржа, %", min_value=0.0, value=20.0, step=1.0)
            return_rate = st.number_input("Доля возвратов, %", min_value=0.0, value=8.0, step=0.5)
            cancel_rate = st.number_input("Доля невыкупа/отмен, %", min_value=0.0, value=3.0, step=0.5)

        if st.button("Рассчитать", type="primary", use_container_width=True):
            source_df = pd.DataFrame([{
                "Артикул": sku,
                "Наименование товара": name,
                "Себестоимость, ₽": cost,
                "Цена продажи, ₽": price,
                "Вес факт, кг": weight,
                "Длина, см": length,
                "Ширина, см": width,
                "Высота, см": height,
                "Тарифная группа": manual_tariff,
                "Товарная группа 3 уровня": manual_lvl3,
                "Дней хранения FBSM": storage_days,
                "Реклама, %": ads_pct,
                "Налог, %": tax_pct,
                "Прочие расходы, ₽": other_costs,
                "Целевая маржа, %": target_margin,
                "Доля возвратов, %": return_rate,
                "Доля невыкупа/отмен, %": cancel_rate,
            }])
            result = calculate_row(
                source_df.iloc[0], reference_df, scheme,
                fbs_weight_basis_map[fbs_weight_basis], fbs_logistics_profile, fbsm_logistics_profile,
                include_fbsm_return_logistics, include_fbsm_defect_handling, include_fbsm_excess_handling,
            )
            render_single_result(result)

    if page == "Пакетный расчет":
        st.header("Пакетный расчет и выгрузка Excel")
        products_file = st.file_uploader("Загрузите Excel с товарами", type=["xlsx"])
        if products_file is None:
            st.info("Загрузите файл товаров по шаблону. Шаблон можно скачать в боковой панели.")
            return
        try:
            products_df = prepare_products(pd.read_excel(products_file))
        except Exception as e:
            st.error(f"Ошибка в файле товаров: {e}")
            st.stop()

        result_rows = []
        for _, row in products_df.iterrows():
            result_rows.append(
                calculate_row(
                    row, reference_df, scheme,
                    fbs_weight_basis_map[fbs_weight_basis], fbs_logistics_profile, fbsm_logistics_profile,
                    include_fbsm_return_logistics, include_fbsm_defect_handling, include_fbsm_excess_handling,
                )
            )
        result_df = pd.DataFrame(result_rows)
        st.dataframe(result_df, use_container_width=True, height=550)

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("SKU", len(result_df))
        m2.metric("Средняя маржа", f"{result_df['Маржа к выручке, %'].mean():.1%}")
        m3.metric("Суммарная прибыль", f"{result_df['Прибыль, ₽'].sum():,.2f} ₽")
        m4.metric("Средняя выплата до себестоимости", f"{result_df['Выплата до себестоимости, ₽'].mean():,.2f} ₽")

        export_bytes = build_export_workbook(result_df)
        st.download_button(
            "Скачать Excel-выгрузку",
            data=export_bytes,
            file_name=f"Спортмастер_юнит-экономика_{scheme}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


if __name__ == "__main__":
    app()

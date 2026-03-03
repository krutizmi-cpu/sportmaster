import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
import requests
from io import BytesIO

# Настройка страницы
st.set_page_config(
    page_title="Спортмастер - FBS Калькулятор",
    page_icon="🏃",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Заголовок приложения
st.title("🏃 Спортмастер - FBS Калькулятор")
st.markdown("Система расчёта юнит-экономики для FBS")

# Боковая панель с навигацией
st.sidebar.title("Навигация")
page = st.sidebar.radio(
    "Выберите раздел:",
    ["📊 Калькулятор", "📈 Аналитика", "📋 История", "⚙️ Настройки"]
)

# Главная страница - Калькулятор
if page == "📊 Калькулятор":
    st.header("Калькулятор юнит-экономики FBS")
    
    # Создаем три колонки для ввода данных
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.subheader("Основные параметры")
        sku = st.text_input("SKU товара", placeholder="Введите артикул")
        name = st.text_input("Название товара", placeholder="Название")
        cost = st.number_input("Себестоимость, ₽", min_value=0.0, step=100.0)
        price = st.number_input("Цена продажи, ₽", min_value=0.0, step=100.0)
        
    with col2:
        st.subheader("Логистика и хранение")
        weight = st.number_input("Вес товара, кг", min_value=0.0, step=0.1)
        volume = st.number_input("Объём, л", min_value=0.0, step=1.0)
        storage_days = st.number_input("Дней на складе", min_value=1, step=1, value=30)
        logistics_cost = st.number_input("Стоимость логистики, ₽", min_value=0.0, step=10.0)
        
    with col3:
        st.subheader("Комиссии и налоги")
        commission_rate = st.slider("Комиссия маркетплейса, %", 0.0, 30.0, 15.0, 0.5)
        tax_rate = st.slider("Налоговая ставка, %", 0.0, 20.0, 6.0, 0.5)
        extra_costs = st.number_input("Доп. расходы, ₽", min_value=0.0, step=50.0)
    
    # Кнопка расчёта
    if st.button("🧮 Рассчитать", type="primary", use_container_width=True):
        if cost > 0 and price > 0:
            # Расчёты
            commission = price * (commission_rate / 100)
            tax = (price - commission) * (tax_rate / 100)
            storage_cost = storage_days * 0.5  # Условная стоимость хранения
            total_costs = cost + logistics_cost + commission + tax + storage_cost + extra_costs
            profit = price - total_costs
            margin = (profit / price * 100) if price > 0 else 0
            roi = (profit / cost * 100) if cost > 0 else 0
            
            # Вывод результатов
            st.success("✅ Расчёт выполнен успешно!")
            
            # Метрики в 4 колонках
            metric1, metric2, metric3, metric4 = st.columns(4)
            
            with metric1:
                st.metric("💰 Прибыль", f"{profit:.2f} ₽")
            with metric2:
                st.metric("📊 Маржинальность", f"{margin:.1f}%")
            with metric3:
                st.metric("📈 ROI", f"{roi:.1f}%")
            with metric4:
                st.metric("💸 Общие расходы", f"{total_costs:.2f} ₽")
            
            # Детализация расходов
            st.subheader("Детализация расходов")
            expense_data = {
                "Статья расходов": ["Себестоимость", "Комиссия МП", "Логистика", "Хранение", "Налоги", "Доп. расходы"],
                "Сумма, ₽": [cost, commission, logistics_cost, storage_cost, tax, extra_costs],
                "Доля, %": [
                    (cost/total_costs*100) if total_costs > 0 else 0,
                    (commission/total_costs*100) if total_costs > 0 else 0,
                    (logistics_cost/total_costs*100) if total_costs > 0 else 0,
                    (storage_cost/total_costs*100) if total_costs > 0 else 0,
                    (tax/total_costs*100) if total_costs > 0 else 0,
                    (extra_costs/total_costs*100) if total_costs > 0 else 0
                ]
            }
            df_expenses = pd.DataFrame(expense_data)
            st.dataframe(df_expenses, use_container_width=True)
            
            # График расходов
            st.bar_chart(df_expenses.set_index("Статья расходов")["Сумма, ₽"])
            
        else:
            st.error("⚠️ Заполните обязательные поля: себестоимость и цену продажи")

# Страница аналитики
elif page == "📈 Аналитика":
    st.header("Аналитика и отчёты")
    st.info("🚧 Раздел в разработке. Здесь будет отображаться аналитика по всем расчётам.")
    
    # Демо данные для примера
    demo_data = {
        "SKU": ["SKU001", "SKU002", "SKU003"],
        "Название": ["Кроссовки Nike", "Мяч футбольный", "Гантели 5кг"],
        "Прибыль, ₽": [1500, 800, 600],
        "Маржа, %": [25, 20, 15],
        "ROI, %": [45, 35, 25]
    }
    st.dataframe(pd.DataFrame(demo_data), use_container_width=True)

# Страница истории
elif page == "📋 История":
    st.header("История расчётов")
    st.info("🚧 Раздел в разработке. Здесь будет храниться история всех ваших расчётов.")
    st.write("Функционал включает:")
    st.markdown("""
    - Сохранение всех расчётов
    - Экспорт в Excel/CSV
    - Фильтрация по датам
    - Сравнение товаров
    """)

# Страница настроек
elif page == "⚙️ Настройки":
    st.header("Настройки приложения")
    st.info("🚧 Раздел в разработке")
    
    st.subheader("Общие настройки")
    currency = st.selectbox("Валюта", ["₽ RUB", "$ USD", "€ EUR"])
    theme = st.selectbox("Тема оформления", ["Светлая", "Тёмная", "Авто"])
    
    st.subheader("Параметры по умолчанию")
    default_commission = st.slider("Комиссия по умолчанию, %", 0.0, 30.0, 15.0, 0.5)
    default_tax = st.slider("Налог по умолчанию, %", 0.0, 20.0, 6.0, 0.5)
    
    if st.button("💾 Сохранить настройки"):
        st.success("✅ Настройки сохранены!")

# Футер
st.sidebar.markdown("---")
st.sidebar.markdown("### 📞 Поддержка")
st.sidebar.markdown("[Документация](#) | [GitHub](#) | [Telegram](#)")
st.sidebar.caption(f"Версия 1.0.0 | © {datetime.now().year}")

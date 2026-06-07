import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import os
import warnings
from io import BytesIO
from datetime import datetime

warnings.filterwarnings('ignore')

# ---------- НАСТРОЙКА СТРАНИЦЫ ----------
st.set_page_config(page_title="ЕИ | Аналитика",
                   page_icon="🏭", layout="wide")
st.title("🏭 Интерактивный дашборд")
st.markdown("""
    <style>
        div.block-container {padding-top: 2rem;}
        .stDataFrame {border-radius: 10px;}
        .css-1d391kg {background-color: #f5f5f5;}
    </style>
""", unsafe_allow_html=True)

# ---------- ФУНКЦИИ ЭКСПОРТА ----------


@st.cache_data
def to_excel_bytes(df, sheet_name="Sheet1"):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return output.getvalue()


# ---------- ЗАГРУЗКА ФАЙЛА И ВЫБОР ЛИСТА ----------
st.subheader("📂 Загрузка данных")
uploaded_file = st.file_uploader(
    "Выберите Excel-файл (.xlsx, .xls, .xlsm)", type=["xlsx", "xls", "xlsm"])
df_raw = None
sheet_names = []
selected_sheet = None

if uploaded_file is not None:
    try:
        xl = pd.ExcelFile(uploaded_file)
        sheet_names = xl.sheet_names
        selected_sheet = st.selectbox("📑 Выберите лист в файле", sheet_names)
        with st.spinner("Чтение данных..."):
            df_raw = pd.read_excel(
                uploaded_file, sheet_name=selected_sheet, engine='openpyxl')
        st.success(
            f"✅ Лист **{selected_sheet}** загружен. Строк: {df_raw.shape[0]}, колонок: {df_raw.shape[1]}")
    except Exception as e:
        st.error(f"Ошибка при загрузке файла: {e}")
else:
    # Путь к файлу по умолчанию (измените при необходимости)
    default_path = r"C:\Users\Alex\Desktop\П26.677.0159-T3-0001\02_ПТ_П26_677_0159_П26_677_0159-T3-0001_v00.xlsx"
    if os.path.exists(default_path):
        try:
            xl = pd.ExcelFile(default_path)
            sheet_names = xl.sheet_names
            selected_sheet = st.selectbox(
                "📑 Выберите лист (файл по умолчанию)", sheet_names)
            with st.spinner("Чтение данных..."):
                df_raw = pd.read_excel(
                    default_path, sheet_name=selected_sheet, engine='openpyxl')
            st.info(
                f"📁 Используется файл по умолчанию, лист **{selected_sheet}**")
        except Exception as e:
            st.error(f"Ошибка при чтении файла по умолчанию: {e}")
            st.stop()
    else:
        st.warning("Пожалуйста, загрузите Excel-файл.")
        st.stop()

if df_raw is None:
    st.stop()

# ---------- ПРЕДОБРАБОТКА: ПЕРЕИМЕНОВАНИЕ КЛЮЧЕВЫХ КОЛОНОК ----------
df = df_raw.copy()
rename_dict = {
    "Начало Прогноз / Факт": "start_date",
    "Окончание Прогноз / Факт": "end_date",
    "Трудозатраты, чел/часы": "ч/ч"
}
for old, new in rename_dict.items():
    if old in df.columns:
        df.rename(columns={old: new}, inplace=True)

# Проверка наличия дат
has_dates = ("start_date" in df.columns and "end_date" in df.columns)
if has_dates:
    df["start_date"] = pd.to_datetime(
        df["start_date"], errors='coerce', dayfirst=True)
    df["end_date"] = pd.to_datetime(
        df["end_date"], errors='coerce', dayfirst=True)
    # Удаляем строки, где обе даты отсутствуют
    df = df.dropna(subset=["start_date", "end_date"], how='all')
    if df.empty:
        st.warning("После обработки дат не осталось строк. Проверьте формат дат.")
        st.stop()
else:
    st.info("ℹ️ В листе не найдены колонки 'Начало Прогноз / Факт' и 'Окончание Прогноз / Факт'. Временная ось недоступна, будут доступны только категориальные графики.")

# ---------- ФИЛЬТРАЦИЯ ПО ДАТАМ (если есть) ----------
df_filtered = df.copy()
if has_dates and not df.empty:
    all_dates = pd.concat([df["start_date"], df["end_date"]]).dropna()
    if not all_dates.empty:
        min_date = all_dates.min()
        max_date = all_dates.max()
        col1, col2 = st.columns(2)
        with col1:
            start_date_widget = st.date_input(
                "📅 Дата начала периода", min_date)
        with col2:
            end_date_widget = st.date_input(
                "📅 Дата окончания периода", max_date)
        start_date_widget = pd.to_datetime(start_date_widget)
        end_date_widget = pd.to_datetime(end_date_widget)
        df_filtered = df[(df["start_date"] <= end_date_widget) & (
            df["end_date"] >= start_date_widget)].copy()

# ---------- БОКОВАЯ ПАНЕЛЬ: ТОЛЬКО НУЖНЫЕ ТЕКСТОВЫЕ ФИЛЬТРЫ ----------
st.sidebar.header("🔍 Фильтры данных")

# Список колонок, которые мы разрешаем для фильтрации (исключаем числа, даты, статусы, стоимости)
allowed_filter_cols = []
# Приоритетные колонки (всегда показываем, если есть)
priority = ["Арх.№", "Наименованиеработы:",
            "Отдел", "Ведущий / Смежный", "ГИП"]
for col in priority:
    if col in df_filtered.columns:
        allowed_filter_cols.append(col)

# Дополнительные текстовые колонки (не содержащие запрещённых слов)
forbidden_words = ["стоимость", "фриланс", "спо",
                   "статус", "дата", "Стоимость", "Статус", "СПО"]
for col in df_filtered.columns:
    # Пропускаем, если имя колонки не строка (например, datetime)
    if not isinstance(col, str):
        continue
    if col in allowed_filter_cols:
        continue
    if col in ["start_date", "end_date", "ч/ч"]:
        continue
    # Только object или category
    if not (pd.api.types.is_object_dtype(df_filtered[col]) or pd.api.types.is_categorical_dtype(df_filtered[col])):
        continue
    # Исключаем по ключевым словам (теперь col точно строка)
    if any(word in col for word in forbidden_words):
        continue
    # Ограничиваем количество уникальных (не более 100)
    if df_filtered[col].nunique() <= 100:
        allowed_filter_cols.append(col)

# Если не осталось ни одной колонки – используем хотя бы приоритетные
if not allowed_filter_cols:
    allowed_filter_cols = [c for c in priority if c in df_filtered.columns]

# Кнопка сброса всех фильтров (настройка состояния)
if "filter_values" not in st.session_state:
    st.session_state.filter_values = {}

# Создаём виджеты фильтров
filter_selections = {}
for col in allowed_filter_cols:
    unique_vals = sorted(df_filtered[col].dropna().astype(str).unique())
    if not unique_vals:
        continue
    # Используем session_state для хранения выбранных значений
    key = f"filter_{col}"
    default = st.session_state.filter_values.get(key, [])
    selected = st.sidebar.multiselect(
        f"📌 {col}", unique_vals, default=default, key=key)
    filter_selections[col] = selected
    # Сохраняем в session_state
    st.session_state.filter_values[key] = selected

# Кнопка сброса всех фильтров
if st.sidebar.button("🧹 Сбросить все фильтры"):
    for col in allowed_filter_cols:
        st.session_state.filter_values[f"filter_{col}"] = []
    st.rerun()

# Применяем фильтры к df_filtered
for col, selected_vals in filter_selections.items():
    if selected_vals:
        df_filtered = df_filtered[df_filtered[col].astype(
            str).isin(selected_vals)]

# ---------- ГЛОБАЛЬНЫЙ ТЕКСТОВЫЙ ПОИСК ----------
st.sidebar.markdown("---")
search_term = st.sidebar.text_input(
    "🔎 Поиск по всем данным (текст / цифры)", placeholder="например, 170 или УПП")
if search_term:
    mask = df_filtered.astype(str).apply(lambda col: col.str.contains(
        search_term, case=False, na=False)).any(axis=1)
    df_filtered = df_filtered[mask]

# ---------- ОТОБРАЖЕНИЕ ТАБЛИЦЫ ----------
st.metric("📊 Отфильтровано строк", len(df_filtered))
st.dataframe(df_filtered, use_container_width=True, height=400)

# ---------- БЫСТРЫЕ ГРАФИКИ (если есть данные) ----------
if not df_filtered.empty:
    st.subheader("📈 Быстрые визуализации")
    if "ч/ч" in df_filtered.columns:
        col1, col2 = st.columns(2)
        # Гистограмма дат (если есть)
        if has_dates and "start_date" in df_filtered.columns:
            with col1:
                fig_hist = px.histogram(
                    df_filtered, x="start_date", title="Распределение дат начала работ", nbins=20)
                st.plotly_chart(fig_hist, use_container_width=True)
        # Топ-10 видов работ по часам
        if "Наименованиеработы:" in df_filtered.columns:
            top_works = df_filtered.groupby("Наименованиеработы:", as_index=False)[
                "ч/ч"].sum().sort_values("ч/ч", ascending=False).head(10)
            if not top_works.empty:
                with col2:
                    fig_bar = px.bar(top_works, x="ч/ч", y="Наименованиеработы:",
                                     orientation='h', title="Топ-10 видов работ (чел·ч)")
                    st.plotly_chart(fig_bar, use_container_width=True)
        # Круговая диаграмма по отделам
        if "Отдел" in df_filtered.columns:
            dept_data = df_filtered.groupby(
                "Отдел", as_index=False)["ч/ч"].sum()
            dept_data = dept_data[dept_data["ч/ч"] > 0]
            if not dept_data.empty:
                fig_pie = px.pie(dept_data, values="ч/ч", names="Отдел",
                                 title="Трудозатраты по отделам", hole=0.4)
                fig_pie.update_traces(
                    textposition="outside", textinfo="percent+label")
                st.plotly_chart(fig_pie, use_container_width=True)

# ========== КОНСТРУКТОР КОМБИНИРОВАННОЙ ДИАГРАММЫ ==========
st.markdown("---")
st.subheader("🎛️ Конструктор комбинированной диаграммы")

if df_filtered.empty:
    st.info("Нет данных для построения диаграммы. Измените фильтры.")
else:
    # 1. Тип оси X
    axis_type = st.radio("Выберите тип оси X:", [
                         "📅 Временная шкала", "🏷️ Категориальное поле"], horizontal=True)
    if axis_type.startswith("📅"):
        if not has_dates:
            st.warning(
                "В данных нет колонок с датами. Переключитесь на категориальную ось.")
            axis_type = "🏷️ Категориальное поле"
            use_time = False
        else:
            period = st.selectbox("Период группировки:", [
                                  "Месяц", "Квартал", "Год"], index=0)
            period_map = {"Месяц": "ME", "Квартал": "QE", "Год": "YE"}
            use_time = True
    if axis_type.startswith("🏷️"):
        # Определяем подходящие категориальные колонки (не даты, не числа)
        cat_cols = []
        for col in df_filtered.columns:
            # Проверяем, что имя колонки – строка
            if not isinstance(col, str):
                continue
            if col in ["start_date", "end_date", "ч/ч"]:
                continue
            if pd.api.types.is_object_dtype(df_filtered[col]) or pd.api.types.is_categorical_dtype(df_filtered[col]):
                if df_filtered[col].nunique() <= 100:
                    cat_cols.append(col)
        if not cat_cols:
            cat_cols = [c for c in ["Отдел", "Наименованиеработы:",
                                    "Арх.№"] if c in df_filtered.columns]
        if not cat_cols:
            st.error(
                "Нет категориальных колонок для оси X. Добавьте текстовые колонки в файл.")
            st.stop()
        category_col = st.selectbox("Категориальное поле:", cat_cols)
        use_time = False

    # 2. Выбор метрик
    numeric_cols = [col for col in df_filtered.columns if col not in [
        "start_date", "end_date"] and pd.api.types.is_numeric_dtype(df_filtered[col])]
    metric_options = numeric_cols + ["📊 Количество работ"]
    if "ч/ч" not in metric_options and "ч/ч" in df_filtered.columns:
        metric_options.append("ч/ч")
    if not metric_options:
        st.warning("Нет числовых колонок для построения графика.")
        st.stop()
    selected_metrics = st.multiselect("📈 Выберите метрики (от 1 до 5 для наглядности):",
                                      metric_options, default=metric_options[:2] if len(metric_options) >= 2 else metric_options)
    if not selected_metrics:
        st.info("Выберите хотя бы одну метрику.")
        st.stop()

    # 3. Настройка типов и осей для каждой метрики
    st.markdown("#### ⚙️ Настройки рядов")
    config = {}
    for i, m in enumerate(selected_metrics):
        c1, c2, c3 = st.columns([2, 1.5, 1.5])
        with c1:
            st.markdown(f"**{m}**")
        with c2:
            chart_type = st.selectbox(
                "Тип", ["Столбчатая", "Линейная", "Область"], key=f"type_{i}", index=0)
        with c3:
            axis = st.selectbox(
                "Ось", ["Левая", "Правая"], key=f"axis_{i}", index=0)
        config[m] = {"type": chart_type, "axis": axis}

    # 4. Агрегация данных
    if use_time:
        ts_data = df_filtered.copy()
        ts_data["start_date"] = pd.to_datetime(ts_data["start_date"])
        groups = ts_data.groupby(pd.Grouper(
            key="start_date", freq=period_map[period]))
        agg_dict = {}
        for m in selected_metrics:
            if m == "📊 Количество работ":
                agg_dict["count"] = ("start_date", "count")
            else:
                agg_dict[m] = (m, "sum")
        agg_df = groups.agg(
            **{k: v for k, v in agg_dict.items()}).reset_index()
        agg_df.rename(columns={"start_date": "period"}, inplace=True)
        if "📊 Количество работ" in selected_metrics:
            agg_df["📊 Количество работ"] = agg_df["count"]
        # Форматируем метку оси
        if period == "Месяц":
            agg_df["x_label"] = agg_df["period"].dt.strftime("%Y-%m")
        elif period == "Квартал":
            agg_df["x_label"] = agg_df["period"].dt.strftime("%Y-Q%q")
        else:
            agg_df["x_label"] = agg_df["period"].dt.strftime("%Y")
        x_vals = agg_df["x_label"]
        x_title = period
        title_suffix = f"по {period.lower()}ам"
    else:
        groups = df_filtered.groupby(category_col)
        agg_dict = {}
        for m in selected_metrics:
            if m == "📊 Количество работ":
                agg_dict["count"] = (category_col, "count")
            else:
                agg_dict[m] = (m, "sum")
        agg_df = groups.agg(
            **{k: v for k, v in agg_dict.items()}).reset_index()
        agg_df.rename(columns={category_col: "category"}, inplace=True)
        if "📊 Количество работ" in selected_metrics:
            agg_df["📊 Количество работ"] = agg_df["count"]
        # Сортировка по первой метрике
        first_metric = selected_metrics[0]
        if first_metric in agg_df.columns:
            agg_df = agg_df.sort_values(first_metric, ascending=False)
        x_vals = agg_df["category"]
        x_title = category_col
        title_suffix = f"по {category_col}"

    if agg_df.empty:
        st.warning(
            "После агрегации данных не осталось записей. Измените параметры.")
    else:
        # 5. Построение графика
        fig = go.Figure()
        # Сначала левая ось, потом правая
        for axis_side in ["Левая", "Правая"]:
            for m in selected_metrics:
                if m not in agg_df.columns:
                    continue
                if config[m]["axis"] != axis_side:
                    continue
                y_vals = agg_df[m]
                name = m
                chart_type = config[m]["type"]
                yaxis_ref = "y" if axis_side == "Левая" else "y2"
                if chart_type == "Столбчатая":
                    fig.add_trace(go.Bar(x=x_vals, y=y_vals,
                                  name=name, yaxis=yaxis_ref))
                elif chart_type == "Линейная":
                    fig.add_trace(go.Scatter(
                        x=x_vals, y=y_vals, mode='lines+markers', name=name, yaxis=yaxis_ref))
                else:  # Область
                    fig.add_trace(go.Scatter(
                        x=x_vals, y=y_vals, mode='lines', fill='tozeroy', name=name, yaxis=yaxis_ref))
        # Настройка макета
        fig.update_layout(
            title=f"📊 Комбинированная диаграмма: {', '.join(selected_metrics)} {title_suffix}",
            xaxis_title=x_title,
            yaxis=dict(title="Значения", side="left"),
            legend=dict(orientation="h", yanchor="bottom",
                        y=1.02, xanchor="right", x=1),
            height=600,
            hovermode="x unified"
        )
        if any(config[m]["axis"] == "Правая" for m in selected_metrics):
            fig.update_layout(yaxis2=dict(
                title="Значения (правая ось)", overlaying="y", side="right", showgrid=False))
        else:
            fig.update_layout(yaxis2=dict(visible=False))
        st.plotly_chart(fig, use_container_width=True)

        # 6. Экспорт агрегированных данных
        with st.expander("📥 Скачать данные для этой диаграммы"):
            if use_time:
                export_df = agg_df[["x_label"] + selected_metrics].copy()
                export_df.rename(columns={"x_label": x_title}, inplace=True)
            else:
                export_df = agg_df[["category"] + selected_metrics].copy()
                export_df.rename(columns={"category": x_title}, inplace=True)
            st.write(export_df.style.background_gradient(cmap="Blues"))
            excel_data = to_excel_bytes(export_df, "CombinedData")
            st.download_button("💾 Скачать Excel", excel_data, f"combined_{'time' if use_time else x_title}.xlsx",
                               "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# ---------- ЭКСПОРТ ВСЕХ ОТФИЛЬТРОВАННЫХ ДАННЫХ ----------
st.markdown("---")
st.subheader("📎 Экспорт данных")
if not df_filtered.empty:
    st.download_button("📥 Скачать отфильтрованную таблицу (Excel)", to_excel_bytes(df_filtered),
                       "filtered_data.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
else:
    st.info("Нет данных для экспорта.")

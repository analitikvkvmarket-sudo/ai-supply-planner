"""Standalone AI supply plan generator. GitHub stores persistent, approved rules."""
from datetime import date
from io import BytesIO
import base64
import json
from pathlib import Path
import requests
import pandas as pd
import streamlit as st
import psycopg

from engine import plan_for_date, compute_matrix, excel_matrix

st.set_page_config(page_title='ИИ — План поставок', layout='wide', page_icon='📦')
st.title('ИИ — План поставок')
st.caption('Готовая матрица SKU × точки → один Excel. Данные продаж и меню — PostgreSQL, модель и правила — GitHub.')

RULES_PATH = 'rules/supply_rules.json'
LOCAL_RULES = Path(__file__).parent / RULES_PATH


def github_config():
    return (st.secrets.get('github_repo', ''), st.secrets.get('github_token', ''))


def rules_load():
    repo, token = github_config()
    if not repo or not token:
        return json.loads(LOCAL_RULES.read_text(encoding='utf-8')), None, False
    url = f'https://api.github.com/repos/{repo}/contents/{RULES_PATH}'
    response = requests.get(url, headers={'Authorization': f'Bearer {token}', 'Accept': 'application/vnd.github+json'}, timeout=15)
    response.raise_for_status()
    doc = response.json()
    return json.loads(base64.b64decode(doc['content']).decode('utf-8')), doc['sha'], True


def rules_save(rules, sha):
    repo, token = github_config()
    if not (repo and token and sha):
        raise ValueError('Для сохранения правил в GitHub настройте github_repo и github_token в Streamlit Secrets.')
    payload = base64.b64encode((json.dumps(rules, ensure_ascii=False, indent=2) + '\n').encode('utf-8')).decode('ascii')
    r = requests.put(f'https://api.github.com/repos/{repo}/contents/{RULES_PATH}',
        headers={'Authorization': f'Bearer {token}', 'Accept': 'application/vnd.github+json'},
        json={'message': 'rules: approved supply exclusion', 'content': payload, 'sha': sha}, timeout=20)
    r.raise_for_status()


def pg_connection():
    dsn = st.secrets.get('postgres_dsn', '')
    if not dsn:
        raise ValueError('Укажите postgres_dsn в Streamlit Secrets, не в GitHub.')
    return psycopg.connect(dsn)


def apply_github_rules(df, rules):
    out = df.copy()
    for rule in rules:
        if not rule.get('active', True) or rule.get('action') != 'exclude':
            continue
        point = str(rule.get('point', '')).replace('T', 'Т')
        scope = rule.get('scope')
        term = str(rule.get('pattern', '')).strip().casefold()
        if not term:
            continue
        match = out['Точка'].eq(point)
        if scope == 'sku':
            match &= out['SKU'].str.lstrip('0').eq(term.lstrip('0'))
        elif scope == 'category':
            match &= out['Категория'].str.casefold().eq(term)
        elif scope == 'name_contains':
            match &= out['product_name'].fillna('').str.casefold().str.contains(term, regex=False)
        else:
            continue
        out.loc[match, 'Рекомендуется'] = 0
        out.loc[match, 'Причина'] = 'Подтверждённый запрет из GitHub'
    out['Изменение'] = out['Рекомендуется'] - out['Текущий план']
    return out

try:
    rules, sha, online_rules = rules_load()
except Exception as exc:
    st.error(f'Не удалось прочитать правила GitHub: {exc}')
    st.stop()

menu_date = st.date_input('Дата меню', value=date.today())
cap = st.slider('Максимальное изменение относительно текущего плана, % (пилот)', 10, 50, 30, step=5)
st.warning('Пилот: CatBoost прогнозирует наблюдаемые продажи, а не гарантированный спрос. Время распродажи, точные остатки и цветовые минимумы ещё не учтены. Проверяй файл перед поставкой.')
if st.button('Сформировать план ИИ', type='primary'):
    try:
        plan = plan_for_date(pg_connection, menu_date)
        if plan.empty:
            st.warning('Для выбранной даты меню в PostgreSQL нет плана.')
            st.session_state.pop('ready', None)
        else:
            result = compute_matrix(plan, menu_date, cap)
            result = apply_github_rules(result, rules)
            result['Утверждаемое количество'] = result['Рекомендуется'].astype(int)
            st.session_state['ready'] = (menu_date, cap, result)
    except Exception as exc:
        st.error(f'Ошибка формирования плана: {exc}')

current = st.session_state.get('ready')
if current and current[0] == menu_date and current[1] == cap:
    result = current[2].copy()
    st.metric('Всего рекомендовано, шт.', int(result['Рекомендуется'].sum()))
    pts = [f'Т{i}' for i in range(1, 41)]
    grid = result.pivot_table(index=['SKU','product_name'], columns='Точка', values='Рекомендуется', aggfunc='first', fill_value=0).reindex(columns=pts, fill_value=0).reset_index()
    grid['ИТОГО'] = grid[pts].sum(axis=1)
    st.dataframe(grid, hide_index=True, use_container_width=True)
    with st.expander('Ручная корректировка перед выгрузкой'):
        fields = ['Точка','SKU','product_name','Категория','Текущий план','Прогноз продаж','Рекомендуется','Утверждаемое количество']
        edit = st.data_editor(result[fields], hide_index=True, disabled=[x for x in fields if x != 'Утверждаемое количество'], key=f'editor_{menu_date}_{cap}', use_container_width=True,
            column_config={'Утверждаемое количество':st.column_config.NumberColumn(min_value=0, step=1, required=True)})
        result['Утверждаемое количество'] = pd.to_numeric(edit['Утверждаемое количество'], errors='coerce').fillna(result['Рекомендуется']).clip(lower=0).round().astype(int)
    st.download_button('Скачать готовый Excel (.xlsx)', data=excel_matrix(result, menu_date), file_name=f'Plan_II_{menu_date}.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', type='primary')

st.divider()
st.subheader('Память ИИ — постоянные запреты')
st.caption('Правила сохраняются версией файла в GitHub. Они применяются к каждому следующему плану. Запреты не обучают CatBoost автоматически.')
if rules:
    st.dataframe(pd.DataFrame(rules), hide_index=True, use_container_width=True)
else:
    st.info('Пока нет сохранённых правил.')
with st.form('rule_form'):
    point = st.selectbox('Точка', [f'Т{i}' for i in range(1, 41)])
    scope = st.selectbox('На что действует запрет', ['name_contains','sku','category'], format_func=lambda x: {'name_contains':'Часть названия (например, каша)','sku':'Конкретный SKU','category':'Категория'}[x])
    pattern = st.text_input('Название / SKU / категория', placeholder='каша')
    confirm = st.checkbox('Подтверждаю постоянный запрет')
    submitted = st.form_submit_button('Запомнить правило в GitHub')
if submitted:
    if not confirm or not pattern.strip():
        st.warning('Заполни правило и подтверди его.')
    elif any(x['point']==point and x['scope']==scope and x['pattern'].casefold()==pattern.strip().casefold() and x.get('active',True) for x in rules):
        st.warning('Такое правило уже есть.')
    else:
        next_rules = rules + [{'point':point, 'scope':scope, 'pattern':pattern.strip(), 'action':'exclude', 'active':True}]
        try:
            rules_save(next_rules, sha)
            st.success('Правило сохранено в GitHub и будет действовать при следующем расчёте.')
            st.rerun()
        except Exception as exc:
            st.error(f'Не удалось записать правило: {exc}')

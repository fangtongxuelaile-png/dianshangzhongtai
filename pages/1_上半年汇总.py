# -*- coding: utf-8 -*-
# 上半年汇总页面 - 仅管理员可访问
from __future__ import annotations
import datetime
import io
import pathlib
import sys
import os

# 添加父目录到路径，以便导入 app.py 中的函数
_sys_path = pathlib.Path(__file__).parent.parent
if str(_sys_path) not in sys.path:
    sys.path.insert(0, str(_sys_path))

import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

# ── 认证检查（管理员 only）──
import hashlib, json as _json_lib

_USERS_FILE = pathlib.Path(__file__).parent.parent / 'users.json'
_SALT = 'xiaotun_bi_2026_salt'

if 'authenticated' not in st.session_state:
    st.session_state.authenticated = False

if not st.session_state.authenticated:
    col_center = st.columns([1, 2, 1])
    with col_center[1]:
        st.markdown('### 🔒 上半年汇总分析', unsafe_allow_html=True)
        st.warning('此页面仅限管理员访问，请先登录')
        with st.form('login_form_h1'):
            _lu = st.text_input('用户名', key='h1_user')
            _lp = st.text_input('密码', type='password', key='h1_pass')
            _submitted = st.form_submit_button('登录', use_container_width=True)
            if _submitted:
                try:
                    with open(_USERS_FILE, 'r', encoding='utf-8') as _f:
                        _users_data = _json_lib.load(_f)
                except Exception:
                    _users_data = {'users': {}}
                _user_info = _users_data.get('users', {}).get(_lu.strip())
                if _user_info:
                    _expire_str = _user_info.get('expire_date', '')
                    if _expire_str:
                        try:
                            _expire_date = datetime.datetime.strptime(_expire_str, '%Y-%m-%d').date()
                            if datetime.date.today() > _expire_date:
                                st.error('⏰ 该账号已过期，请联系管理员')
                                st.stop()
                        except Exception:
                            pass
                    _h = hashlib.sha256((_SALT + _lp).encode()).hexdigest()
                    if _h == _user_info.get('password_hash', ''):
                        st.session_state.authenticated = True
                        st.session_state.username = _lu.strip()
                        st.session_state.role = _user_info.get('role', 'viewer')
                        st.rerun()
                st.error('❌ 用户名或密码错误，或账号未授权')
    st.stop()

# 检查是否为管理员
if st.session_state.role != 'admin':
    st.error('⛔ 此页面仅限管理员访问')
    st.stop()

# ── 页面配置 ──
st.set_page_config(
    page_title='上半年汇总 - 小豚BI',
    page_icon='📊',
    layout='wide',
    initial_sidebar_state='collapsed'
)

# ── CSS 样式（简化版）──
CSS = '''
<style>
.stApp {background: #f8fafc;}
.hero {background: linear-gradient(135deg, #0f172a 0%, #1e293b 100%); border-radius: 16px; padding: 20px 28px; margin: 4px 0 16px 0; color: #fff;}
.hero-title {font-size: 24px; font-weight: 900; margin: 4px 0 4px 0;}
.hero-sub {font-size: 13px; color: #94a3b8; margin-top: 4px;}
/* 表格样式 */
.styled-table-wrap {overflow-x: auto; border-radius: 12px; border: 1px solid #e2e8f0; margin-top: 6px;}
.styled-table {width: 100%; border-collapse: collapse; font-size: 12.5px;}
.styled-table thead th {background: #1e293b; color: #fff; font-weight: 700; text-align: left; padding: 9px 10px; white-space: nowrap;}
.styled-table tbody td {padding: 7px 10px; border-bottom: 1px solid #e5e7eb; vertical-align: middle;}
.styled-table tbody tr:hover {background: #eff6ff;}
</style>
'''
st.markdown(CSS, unsafe_allow_html=True)

st.markdown('''
<div class="hero">
    <div><span style="background:#22c55e20;color:#22c55e;padding:3px 10px;border-radius:20px;font-size:12px;font-weight:600;margin-right:6px;">管理员</span><span style="background:#1d4ed820;color:#1d4ed8;padding:3px 10px;border-radius:20px;font-size:12px;font-weight:600;">上半年汇总</span></div>
    <h1 class="hero-title">📊 2026年上半年经营复盘</h1>
    <div class="hero-sub">华为三GAP方法论 · 全渠道目标达成分析</div>
</div>
''', unsafe_allow_html=True)

# ── 数据加载 ──
@st.cache_data(show_spinner='正在加载数据...', ttl=3600)
def _load_sales_data():
    """加载销售数据"""
    _CACHE_DIR = pathlib.Path(__file__).parent.parent / '.data_cache'
    _REPO_DATA_DIR = pathlib.Path(__file__).parent.parent / 'data'
    _CACHE_SALES = _CACHE_DIR / 'last_sales.xlsx'
    _REPO_SALES = _REPO_DATA_DIR / 'sales.xlsx'
    _CACHE_SALES_PKL = _CACHE_DIR / 'last_sales.pkl'
    
    import pickle
    
    # 尝试pickle缓存
    if _CACHE_SALES_PKL.exists():
        try:
            with open(_CACHE_SALES_PKL, 'rb') as f:
                return pickle.load(f), True
        except Exception:
            pass
    
    # 从文件加载
    file_bytes = None
    if _CACHE_SALES.exists():
        file_bytes = _CACHE_SALES.read_bytes()
    elif _REPO_SALES.exists():
        file_bytes = _REPO_SALES.read_bytes()
    
    if not file_bytes:
        return {}, False
    
    try:
        from dashboard_core import parse_sales_workbook
        result = parse_sales_workbook(io.BytesIO(file_bytes))
        # 保存pickle缓存
        try:
            with open(_CACHE_SALES_PKL, 'wb') as f:
                pickle.dump(result, f)
        except Exception:
            pass
        return result, True
    except Exception as e:
        st.error(f'销售数据加载失败: {e}')
        return {}, False

@st.cache_data(show_spinner='正在加载推广数据...', ttl=3600)
def _load_promo_data():
    """加载推广数据"""
    _CACHE_DIR = pathlib.Path(__file__).parent.parent / '.data_cache'
    _REPO_DATA_DIR = pathlib.Path(__file__).parent.parent / 'data'
    _CACHE_PROMO = _CACHE_DIR / 'last_promo.xlsx'
    _REPO_PROMO = _REPO_DATA_DIR / 'promo.xlsx'
    _CACHE_PROMO_PKL = _CACHE_DIR / 'last_promo.pkl'
    
    import pickle
    
    # 尝试pickle缓存
    if _CACHE_PROMO_PKL.exists():
        try:
            with open(_CACHE_PROMO_PKL, 'rb') as f:
                return pickle.load(f), True
        except Exception:
            pass
    
    # 从文件加载
    file_bytes = None
    if _CACHE_PROMO.exists():
        file_bytes = _CACHE_PROMO.read_bytes()
    elif _REPO_PROMO.exists():
        file_bytes = _REPO_PROMO.read_bytes()
    
    if not file_bytes:
        return [], False
    
    try:
        from dashboard_core import load_promo_data as _load_promo
        result = _load_promo(file_bytes)
        # 保存pickle缓存
        try:
            with open(_CACHE_PROMO_PKL, 'wb') as f:
                pickle.dump(result, f)
        except Exception:
            pass
        return result, True
    except Exception as e:
        st.error(f'推广数据加载失败: {e}')
        return [], False

@st.cache_data(show_spinner='正在加载目标数据...', ttl=3600)
def _load_targets_data():
    """加载目标数据"""
    _CACHE_DIR = pathlib.Path(__file__).parent.parent / '.data_cache'
    _REPO_DATA_DIR = pathlib.Path(__file__).parent.parent / 'data'
    _CACHE_TARGETS = _CACHE_DIR / 'last_targets.xlsx'
    _REPO_TARGETS = _REPO_DATA_DIR / 'targets.xlsx'
    
    # 优先：按月缓存目录
    targets = {}
    _CACHE_TARGETS_DIR = _CACHE_DIR / 'targets'
    if _CACHE_TARGETS_DIR.exists():
        for month_file in sorted(_CACHE_TARGETS_DIR.glob('targets_*.xlsx')):
            try:
                file_bytes = month_file.read_bytes()
                from dashboard_core import load_targets as _load_tgt
                month_targets = _load_tgt(file_bytes)
                targets.update(month_targets)
            except Exception:
                pass
    
    if not targets:
        # 回退：完整文件
        file_bytes = None
        if _CACHE_TARGETS.exists():
            file_bytes = _CACHE_TARGETS.read_bytes()
        elif _REPO_TARGETS.exists():
            file_bytes = _REPO_TARGETS.read_bytes()
        
        if file_bytes:
            try:
                from dashboard_core import load_targets as _load_tgt
                targets = _load_tgt(file_bytes)
            except Exception as e:
                st.error(f'目标数据加载失败: {e}')
    
    return targets

# 加载数据
with st.spinner('正在加载数据...'):
    data, _sales_ok = _load_sales_data()
    promo_rows, _promo_ok = _load_promo_data()
    targets = _load_targets_data()

if not _sales_ok:
    st.warning('⚠️ 未找到销售数据，请在主页面上传数据')
    st.stop()

# ── 上半年汇总视图代码 ──
from data.monthly_targets_1_4 import MONTHLY_TARGETS_1_4, H1_MONTHS

# 辅助函数
_wan = lambda x: f'{x/10000:.1f}万' if abs(x) >= 10000 else f'{x:,.0f}'
_fmt_big = lambda x: f'{x/10000:,.1f}万' if abs(x) >= 10000 else f'{x:,.0f}'

# ── 目标数据（1-4月配置 + 5-6月 targets）──
_targets_5_6 = {}
_aliases_target = ['成交金额目标', '销额目标']
for ym in ['2026-05', '2026-06']:
    if ym in targets:
        for sr in targets[ym].get('shop', []):
            shop = sr['店铺']
            if shop == '天猫小豚':
                continue
            if sr['指标'] in _aliases_target:
                _targets_5_6.setdefault(shop, {})[ym] = sr.get('合计', 0.0) or 0.0

_shop_order = ['华为京东自营', '天猫华为官旗', '天猫智选', '京东小豚', '抖音小豚']
_h1_targets = {}
for shop in _shop_order:
    _h1_targets[shop] = {}
    if shop in MONTHLY_TARGETS_1_4:
        for ym in ['2026-01', '2026-02', '2026-03', '2026-04']:
            _h1_targets[shop][ym] = MONTHLY_TARGETS_1_4[shop].get(ym, 0.0)
    for ym in ['2026-05', '2026-06']:
        _h1_targets[shop][ym] = _targets_5_6.get(shop, {}).get(ym, 0.0)

# ── 实际数据（今年 + 去年同比）──
_month_labels = ['1月', '2月', '3月', '4月', '5月', '6月']
_h1_actual = {shop: {ym: 0.0 for ym in H1_MONTHS} for shop in _shop_order}
_h1_spend = {shop: {ym: 0.0 for ym in H1_MONTHS} for shop in _shop_order}
_yoy_months = ['2025-01', '2025-02', '2025-03', '2025-04', '2025-05', '2025-06']
_h1_actual_ly = {shop: {ym: 0.0 for ym in _yoy_months} for shop in _shop_order}
_h1_spend_ly = {shop: {ym: 0.0 for ym in _yoy_months} for shop in _shop_order}

if 'monthly' in data:
    for r in data['monthly']:
        ym = r.get('月份', '')[:7]
        shop = (r.get('店铺', '') or '').strip()
        amt = float(r.get('支付金额', 0) or 0)
        if ym in H1_MONTHS and shop in _h1_actual:
            _h1_actual[shop][ym] += amt
        elif ym in _yoy_months and shop in _h1_actual_ly:
            _h1_actual_ly[shop][ym] += amt

for r in promo_rows:
    d = r.get('_date', '')
    ym = d[:7] if len(d) >= 7 else ''
    shop = (r.get('_店铺', '') or '').strip()
    spend = float(r.get('_花费', 0) or 0)
    if ym in H1_MONTHS and shop in _h1_spend:
        _h1_spend[shop][ym] += spend
    elif ym in _yoy_months and shop in _h1_spend_ly:
        _h1_spend_ly[shop][ym] += spend

# ── 汇总计算 ──
_total_target = sum(sum(v.values()) for v in _h1_targets.values())
_total_actual = sum(sum(v.values()) for v in _h1_actual.values())
_total_spend = sum(sum(v.values()) for v in _h1_spend.values())
_total_actual_ly = sum(sum(v.values()) for v in _h1_actual_ly.values())
_total_spend_ly = sum(sum(v.values()) for v in _h1_spend_ly.values())

# GAP1: 目标达成
_gap1 = (_total_actual - _total_target) / _total_target if _total_target > 0 else None
# GAP2: 同比（销额）
_gap2 = (_total_actual - _total_actual_ly) / _total_actual_ly if _total_actual_ly > 0 else None
# 费率
_fee_rate = _total_spend / _total_actual * 100 if _total_actual > 0 else 0
_fee_rate_ly = _total_spend_ly / _total_actual_ly * 100 if _total_actual_ly > 0 else 0
_fee_gap = _fee_rate - _fee_rate_ly

# ── 健康评分 ──
_score_parts = []
if _gap1 is not None:
    _s1 = max(0, min(100, 100 + _gap1 * 200))
    _score_parts.append(_s1)
if _gap2 is not None:
    _s2 = max(0, min(100, 100 + _gap2 * 200))
    _score_parts.append(_s2)

_health_score = sum(_score_parts) / len(_score_parts) if _score_parts else 50
if _health_score >= 90:   _status, _st_color, _st_msg = '🟢 健康', '#22c55e', '上半年各项核心指标表现良好，目标与同比均达成预期。'
elif _health_score >= 70: _status, _st_color, _st_msg = '🟡 关注', '#f59e0b', f'部分指标存在差距（综合评分{_health_score:.0f}/100），建议关注标红项。'
elif _health_score >= 50: _status, _st_color, _st_msg = '🔴 风险', '#ef4444', f'多项指标偏离目标（综合评分{_health_score:.0f}/100），需立即采取行动。'
else:                     _status, _st_color, _st_msg = '⚠️ 告警', '#dc2626', f'上半年经营状况堪忧（综合评分{_health_score:.0f}/100），请优先处理P0任务。'

# ════════════════════════════════════════
# Layer 1: 标题 + 一句话总结
# ════════════════════════════════════════
_one_liner = f"上半年全渠道GMV {_fmt_big(_total_actual)}"
if _gap1 is not None:
    _one_liner += f"，目标达成率 {_total_actual/_total_target*100:.1f}%（{'超额' if _gap1 >= 0 else '缺口'}{abs(_gap1)*100:.1f}%）"
if _gap2 is not None:
    _one_liner += f"，同比 {'增长' if _gap2 >= 0 else '下滑'}{abs(_gap2)*100:.1f}%"
_one_liner += f"，实际费率 {_fee_rate:.1f}%"

st.markdown(f"""
<div style='background:linear-gradient(135deg,#0f172a 0%,#1e293b 100%);border-radius:16px;padding:20px 28px;margin:4px 0 16px 0;color:#fff;'>
<div style='font-size:13px;color:#94a3b8;margin-bottom:4px;'>🎯 上半年经营复盘（2026年1-6月）</div>
<div style='font-size:20px;font-weight:800;line-height:1.5;'>{_one_liner}</div>
<div style='display:flex;gap:12px;margin-top:10px;'>
<span style='background:{_st_color}20;color:{_st_color};padding:3px 10px;border-radius:20px;font-size:12px;font-weight:600;'>{_status} · 综合评分 {_health_score:.0f}/100</span>
</div></div>""", unsafe_allow_html=True)

# ════════════════════════════════════════
# Layer 2: 三GAP速览卡片
# ════════════════════════════════════════
st.markdown("<div style='font-size:13px;color:#475569;font-weight:700;margin:4px 0 2px 0;'>📊 三GAP速览</div>", unsafe_allow_html=True)
st.caption('GAP1=目标达成 | GAP2=同比变化 | GAP3=费率变化')

g1, g2, g3, g4 = st.columns(4)

with g1:
    _c = '#22c55e' if (_gap1 or 0) >= 0 else '#ef4444' if (_gap1 or 0) < -0.1 else '#f59e0b'
    _icon = '✅' if (_gap1 or 0) >= 0 else '🔴'
    st.markdown(f"""<div style='background:#f8fafc;border:1px solid #e2e8f0;border-radius:12px;padding:14px;text-align:center;'>
    <div style='font-size:11px;color:#64748b;font-weight:700;'>🎯 GAP1: 目标达成</div>
    <div style='font-size:22px;font-weight:900;color:{_c};'>{_icon} {_gap1*100:+.1f}%</div>
    <div style='font-size:10px;color:#94a3b8;'>实际 {_fmt_big(_total_actual)} / 目标 {_fmt_big(_total_target)}</div></div>""", unsafe_allow_html=True)

with g2:
    _c = '#22c55e' if (_gap2 or 0) >= 0 else '#ef4444' if (_gap2 or 0) < -0.1 else '#f59e0b'
    _icon = '📈' if (_gap2 or 0) >= 0 else '📉'
    st.markdown(f"""<div style='background:#f8fafc;border:1px solid #e2e8f0;border-radius:12px;padding:14px;text-align:center;'>
    <div style='font-size:11px;color:#64748b;font-weight:700;'>📅 GAP2: 同比变化</div>
    <div style='font-size:22px;font-weight:900;color:{_c};'>{_icon} {_gap2*100:+.1f}%</div>
    <div style='font-size:10px;color:#94a3b8;'>今年 {_fmt_big(_total_actual)} / 去年 {_fmt_big(_total_actual_ly)}</div></div>""", unsafe_allow_html=True)

with g3:
    _c = '#22c55e' if _fee_gap <= 0 else '#ef4444' if _fee_gap > 2 else '#f59e0b'
    _icon = '✅' if _fee_gap <= 0 else '⚠️'
    st.markdown(f"""<div style='background:#f8fafc;border:1px solid #e2e8f0;border-radius:12px;padding:14px;text-align:center;'>
    <div style='font-size:11px;color:#64748b;font-weight:700;'>💰 GAP3: 费率变化</div>
    <div style='font-size:22px;font-weight:900;color:{_c};'>{_icon} {_fee_gap:+.1f}pp</div>
    <div style='font-size:10px;color:#94a3b8;'>今年 {_fee_rate:.1f}% / 去年 {_fee_rate_ly:.1f}%</div></div>""", unsafe_allow_html=True)

with g4:
    _c = '#22c55e' if _fee_rate <= 15 else '#ef4444' if _fee_rate > 25 else '#f59e0b'
    st.markdown(f"""<div style='background:#f8fafc;border:1px solid #e2e8f0;border-radius:12px;padding:14px;text-align:center;'>
    <div style='font-size:11px;color:#64748b;font-weight:700;'>📊 总推广花费</div>
    <div style='font-size:22px;font-weight:900;color:#0f172a;'>{_fmt_big(_total_spend)}</div>
    <div style='font-size:10px;color:#94a3b8;'>费率 {_fee_rate:.1f}% | ROI {_total_actual/_total_spend:.1f}</div></div>""", unsafe_allow_html=True)

# 辅助函数：格式化同比
def _fmt_yoy(v):
    if v is None:
        return '<span style="color:#94a3b8;">同比--</span>'
    c = '#22c55e' if v >= 0 else '#ef4444'
    return f'<span style="color:{c};">同比{v:+.0f}%</span>'

# ════════════════════════════════════════
# Layer 3: 各店铺月度达成 + 同比
# ════════════════════════════════════════
st.markdown('---')
st.subheader('📊 各店铺月度达成 & 同比')

_tbl_rows = []
for shop in _shop_order:
    if shop not in _h1_targets:
        continue
    _row_cells = [f'<b>{shop}</b>']
    _st = 0.0; _sa = 0.0; _sal = 0.0; _ssp = 0.0
    for i, ym in enumerate(H1_MONTHS):
        t = _h1_targets[shop].get(ym, 0)
        a = _h1_actual[shop].get(ym, 0)
        _st += t; _sa += a
        ly_ym = _yoy_months[i]
        al = _h1_actual_ly[shop].get(ly_ym, 0)
        _sal += al
        sp = _h1_spend[shop].get(ym, 0)
        _ssp += sp
        rate = a / t * 100 if t > 0 else None
        yoy = (a - al) / al * 100 if al > 0 else None
        if rate is not None:
            rc = '#22c55e' if rate >= 100 else '#ef4444' if rate < 80 else '#f59e0b'
        else:
            rc = '#94a3b8'
        _yoy_str = _fmt_yoy(yoy)
        _cell = f'{_wan(t)} / {_wan(a)}<br><span style="color:{rc};font-weight:600;">{rate:.1f}%</span> {_yoy_str}' if rate is not None else f'{_wan(t)} / {_wan(a)}<br>-- {_yoy_str}'
        _row_cells.append(_cell)
    # 合计
    _sr = _sa / _st * 100 if _st > 0 else None
    _sy = (_sa - _sal) / _sal * 100 if _sal > 0 else None
    if _sr is not None:
        src = '#22c55e' if _sr >= 100 else '#ef4444' if _sr < 80 else '#f59e0b'
    else:
        src = '#94a3b8'

    _syoy_str = _fmt_yoy(_sy)
    _fee = _ssp / _sa * 100 if _sa > 0 else 0
    _row_cells.append(f'<b>{_wan(_st)} / {_wan(_sa)}</b><br><span style="color:{src};font-weight:700;">{_sr:.1f}%</span> {_syoy_str}')
    _tbl_rows.append(_row_cells)

# 全渠道合计
_total_row = ['<b>全渠道合计</b>']
_at = 0.0; _aa = 0.0; _aal = 0.0; _asp = 0.0
for i, ym in enumerate(H1_MONTHS):
    t = sum(_h1_targets[s].get(ym, 0) for s in _shop_order if s in _h1_targets)
    a = sum(_h1_actual[s].get(ym, 0) for s in _shop_order)
    _at += t; _aa += a
    ly_ym = _yoy_months[i]
    al = sum(_h1_actual_ly[s].get(ly_ym, 0) for s in _shop_order)
    _aal += al
    rate = a / t * 100 if t > 0 else None
    yoy = (a - al) / al * 100 if al > 0 else None
    rc = '#22c55e' if (rate or 0) >= 100 else '#ef4444' if (rate or 0) < 80 else '#f59e0b' if rate is not None else '#94a3b8'
    _yoy_str = _fmt_yoy(yoy)
    _total_row.append(f'<b>{_wan(t)} / {_wan(a)}</b><br><span style="color:{rc};font-weight:700;">{rate:.1f}%</span> {_yoy_str}' if rate is not None else f'<b>{_wan(t)} / {_wan(a)}</b><br>-- {_yoy_str}')
_tr = _aa / _at * 100 if _at > 0 else None
_ty = (_aa - _aal) / _aal * 100 if _aal > 0 else None
_trc = '#22c55e' if (_tr or 0) >= 100 else '#ef4444' if (_tr or 0) < 80 else '#f59e0b'
_tyoy_str = _fmt_yoy(_ty)
_total_row.append(f'<b>{_wan(_at)} / {_wan(_aa)}</b><br><span style="color:{_trc};font-weight:700;">{_tr:.1f}%</span> {_tyoy_str}')
_tbl_rows.append(_total_row)

_tbl_id = 'h1_' + str(hash(tuple(H1_MONTHS)))[:8]
_h1_html = f'<div class="styled-table-wrap"><table class="styled-table">'
_h1_html += '<thead><tr><th style="text-align:left;">店铺</th>'
for ml in _month_labels:
    _h1_html += f'<th style="text-align:center;font-size:11px;">{ml}<br><span style="font-weight:400;color:#94a3b8;">目标/实际</span></th>'
_h1_html += '<th style="text-align:center;">上半年合计</th></tr></thead><tbody>'
for _r in _tbl_rows:
    _h1_html += '<tr>'
    for j, cell in enumerate(_r):
        align = 'left' if j == 0 else 'center'
        _h1_html += f'<td style="text-align:{align};white-space:nowrap;font-size:12px;">{cell}</td>'
    _h1_html += '</tr>'
_h1_html += '</tbody></table></div>'
st.markdown(_h1_html, unsafe_allow_html=True)

# ════════════════════════════════════════
# Layer 4: 图表 — 达成率趋势 + 同比
# ════════════════════════════════════════
c_left, c_right = st.columns(2)

with c_left:
    st.markdown('<div style="font-weight:700;font-size:13px;">📈 月度达成率趋势</div>', unsafe_allow_html=True)
    _chart_data = []
    for shop in _shop_order:
        if shop not in _h1_targets:
            continue
        for i, ym in enumerate(H1_MONTHS):
            t = _h1_targets[shop].get(ym, 0)
            a = _h1_actual[shop].get(ym, 0)
            rate = a / t * 100 if t > 0 else None
            _chart_data.append({'店铺': shop, '月份': _month_labels[i], '达成率': rate if rate is not None else 0})
    if _chart_data:
        import pandas as _pd2
        _df_chart = _pd2.DataFrame(_chart_data)
        _fig = px.line(_df_chart, x='月份', y='达成率', color='店铺', markers=True,
                      color_discrete_sequence=['#1d4ed8', '#e6a817', '#22c55e', '#ef4444', '#8b5cf6'])
        _fig.add_hline(y=100, line_dash='dash', line_color='#94a3b8', annotation_text='100%')
        _fig.update_layout(height=340, margin=dict(l=20, r=20, t=10, b=10),
                          hovermode='x unified', yaxis=dict(ticksuffix='%'))
        st.plotly_chart(_fig, use_container_width=True)

with c_right:
    st.markdown('<div style="font-weight:700;font-size:13px;">📊 月度GMV同比变化</div>', unsafe_allow_html=True)
    _yoy_chart = []
    for i, ym in enumerate(H1_MONTHS):
        a = sum(_h1_actual[s].get(ym, 0) for s in _shop_order)
        ly_ym = _yoy_months[i]
        al = sum(_h1_actual_ly[s].get(ly_ym, 0) for s in _shop_order)
        yoy = (a - al) / al * 100 if al > 0 else 0
        _yoy_chart.append({'月份': _month_labels[i], 'GMV(万)': a/10000, '同比': yoy})
    if _yoy_chart:
        _df_yoy = _pd2.DataFrame(_yoy_chart)
        _fig2 = make_subplots(specs=[[{"secondary_y": True}]])
        _fig2.add_trace(go.Bar(x=_df_yoy['月份'], y=_df_yoy['GMV(万)'], name='GMV(万)',
                              marker_color='#1d4ed8', yaxis='y1'), secondary_y=False)
        _fig2.add_trace(go.Scatter(x=_df_yoy['月份'], y=_df_yoy['同比'], name='同比%',
                                  mode='lines+markers', marker_color='#ef4444',
                                  line=dict(width=3)), secondary_y=True)
        _fig2.add_hline(y=0, line_dash='dash', line_color='#94a3b8', secondary_y=True)
        _fig2.update_layout(height=340, margin=dict(l=20, r=20, t=10, b=10),
                           hovermode='x unified', legend=dict(orientation='h', y=1.12))
        _fig2.update_yaxes(title_text='GMV(万)', secondary_y=False)
        _fig2.update_yaxes(title_text='同比%', ticksuffix='%', secondary_y=True)
        st.plotly_chart(_fig2, use_container_width=True)

# ════════════════════════════════════════
# Layer 5: 各店铺累计达成进度条 + 费率
# ════════════════════════════════════════
st.markdown('---')
st.subheader('📊 各店铺累计达成 & 费率')
for shop in _shop_order:
    if shop not in _h1_targets:
        continue
    _s_tgt = sum(_h1_targets[shop].values())
    _s_act = sum(_h1_actual[shop].values())
    _s_spd = sum(_h1_spend[shop].values())
    _s_rate = _s_act / _s_tgt * 100 if _s_tgt > 0 else 0
    _s_fee = _s_spd / _s_act * 100 if _s_act > 0 else 0
    _s_act_ly = sum(_h1_actual_ly[shop].values())
    _s_yoy = (_s_act - _s_act_ly) / _s_act_ly * 100 if _s_act_ly > 0 else None
    _sc = '#22c55e' if _s_rate >= 100 else '#ef4444' if _s_rate < 80 else '#f59e0b'
    _prog = min(_s_rate, 100)
    _yoy_shop = f' | 同比 {_s_yoy:+.1f}%' if _s_yoy is not None else ''
    st.markdown(
        f'<div style="display:flex;align-items:center;margin:4px 0;gap:8px;">'
        f'<span style="width:110px;font-weight:600;font-size:12px;">{shop}</span>'
        f'<div style="flex:1;background:#e5e7eb;border-radius:4px;height:12px;">'
        f'<div style="width:{_prog:.1f}%;background:{_sc};height:12px;border-radius:4px;"></div></div>'
        f'<span style="font-size:11px;color:{_sc};font-weight:600;min-width:70px;text-align:right;">{_s_rate:.1f}%{_yoy_shop}</span>'
        f'<span style="font-size:11px;color:#64748b;min-width:80px;text-align:right;">费率 {_s_fee:.1f}%</span>'
        f'<span style="font-size:10px;color:#94a3b8;min-width:90px;text-align:right;">{_wan(_s_act)} / {_wan(_s_tgt)}</span></div>',
        unsafe_allow_html=True
    )

# ════════════════════════════════════════
# Layer 6: 单品汇总（5-6月）
# ════════════════════════════════════════
st.markdown('---')
st.subheader('📦 单品汇总（5-6月）')
st.caption('仅5-6月有单品级目标拆解数据')

_model_targets = {}
for ym in ['2026-05', '2026-06']:
    if ym not in targets:
        continue
    for mr in targets[ym].get('model', []):
        if mr['指标'] in _aliases_target:
            key = (mr['店铺'], mr['型号'])
            _model_targets.setdefault(key, {})[ym] = mr.get('合计', 0.0) or 0.0

_model_actual = {}
if 'monthly' in data:
    for r in data['monthly']:
        ym = r.get('月份', '')[:7]
        if ym not in ['2026-05', '2026-06']:
            continue
        shop = (r.get('店铺', '') or '').strip()
        model = (r.get('型号', '') or '').strip()
        if not model:
            continue
        key = (shop, model)
        _model_actual.setdefault(key, {'2026-05': 0.0, '2026-06': 0.0})[ym] += float(r.get('支付金额', 0) or 0)

# _MODEL_MERGE_MAP 需要在页面中定义或从app.py导入
_MODEL_MERGE_MAP = {
    '京东小豚': {
        'XT-X60': ['XT-X60【7W】', 'XT-X60【15W】'],
    },
}

for shop_name, shop_map in _MODEL_MERGE_MAP.items():
    for mapped_model, source_models in shop_map.items():
        for ym in ['2026-05', '2026-06']:
            merged = sum(_model_actual.get((shop_name, sm), {}).get(ym, 0) for sm in source_models)
            if merged > 0:
                _model_actual.setdefault((shop_name, mapped_model), {'2026-05': 0.0, '2026-06': 0.0})[ym] += merged

_model_summary = []
for (shop, model), ym_targets in _model_targets.items():
    _mt = sum(ym_targets.values())
    _ma = sum(_model_actual.get((shop, model), {}).values())
    _mr = _ma / _mt * 100 if _mt > 0 else None
    _model_summary.append((shop, model, _mt, _ma, _mr))

_model_summary.sort(key=lambda x: x[2], reverse=True)

if _model_summary:
    _mtbl_rows = []
    for shop, model, mt, ma, mr in _model_summary:
        if mr is not None:
            rc = '#22c55e' if mr >= 100 else '#ef4444' if mr < 80 else '#f59e0b'
            _rate_str = f'<span style="color:{rc};font-weight:600;">{mr:.1f}%</span>'
        else:
            _rate_str = '--'
        _mtbl_rows.append([shop, model, f'{mt:,.0f}', f'{ma:,.0f}', _rate_str])
    _mtbl_html = '<div class="styled-table-wrap"><table class="styled-table">'
    _mtbl_html += '<thead><tr><th>店铺</th><th>型号</th><th>目标销额</th><th>实际销额</th><th>达成率</th></tr></thead><tbody>'
    for _r in _mtbl_rows:
        _mtbl_html += '<tr>'
        for j, cell in enumerate(_r):
            align = 'left' if j <= 1 else 'right'
            _mtbl_html += f'<td style="text-align:{align};white-space:nowrap;">{cell}</td>'
        _mtbl_html += '</tr>'
    _mtbl_html += '</tbody></table></div>'
    st.markdown(_mtbl_html, unsafe_allow_html=True)
else:
    st.info('5-6月暂无单品目标数据')

st.markdown('---')
st.caption('💡 此页面仅管理员可访问，数据每1小时自动刷新缓存')

import streamlit as st
import pandas as pd
import io
import traceback

# ──────────────────────────────────────────────────────────────
# PAGE CONFIG
# ──────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Data Quality Report",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# ──────────────────────────────────────────────────────────────
# CUSTOM CSS
# ──────────────────────────────────────────────────────────────
st.markdown("""
<style>
.main-header {
    background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%);
    color: white; padding: 16px 32px; border-radius: 10px;
    margin-bottom: 20px; display: flex; justify-content: space-between; align-items: center;
}
.main-header h1 { font-size: 22px; font-weight: 600; margin: 0; color: white; }
.main-header .subtitle { font-size: 12px; opacity: 0.7; margin-top: 2px; }
.main-header .hdr-badge { background: rgba(255,255,255,0.12); padding: 6px 14px; border-radius: 20px; font-size: 13px; }
.kpi-row { display: flex; gap: 12px; flex-wrap: wrap; margin-bottom: 18px; }
.kpi-card { flex:1; min-width:130px; background:white; padding:16px 20px; border-radius:10px; text-align:center; box-shadow:0 1px 3px rgba(0,0,0,0.08); border:1px solid #eee; }
.kpi-card .value { font-size: 28px; font-weight: 700; }
.kpi-card .label { font-size:11px; color:#6b778c; text-transform:uppercase; letter-spacing:0.8px; margin-top:4px; }
.kpi-green .value { color:#00875a; } .kpi-red .value { color:#de350b; }
.kpi-orange .value { color:#ff991f; } .kpi-blue .value { color:#0046FF; }
.bg { display:inline-block; padding:3px 10px; border-radius:12px; font-size:12px; font-weight:600; }
.bg-green { background:#e3fcef; color:#006644; } .bg-red { background:#ffebe6; color:#bf2600; }
.bg-orange { background:#fff7e6; color:#974f0c; } .bg-blue { background:#deebff; color:#0747a6; }
.sec-card { background:white; border-radius:10px; box-shadow:0 1px 3px rgba(0,0,0,0.08); padding:20px; margin-bottom:16px; border:1px solid #eee; }
.sec-card h3 { font-size:15px; font-weight:600; margin:0 0 12px 0; color:#172b4d; }
.pct-bar { display:inline-block; height:8px; border-radius:4px; min-width:4px; vertical-align:middle; margin-right:6px; }
.block-container { padding-top: 1rem; max-width: 1400px; }
div[data-testid="stMetric"] { background:white; padding:16px; border-radius:10px; box-shadow:0 1px 3px rgba(0,0,0,0.08); }
div[data-testid="stMetric"] label { font-size:11px !important; text-transform:uppercase; letter-spacing:0.8px; }
#MainMenu {visibility:hidden;} footer {visibility:hidden;} header {visibility:hidden;}
</style>
""", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────
# SAFE STRING HELPERS (avoid .str accessor issues on Py 3.14)
# ──────────────────────────────────────────────────────────────
def safe_str(val):
    """Convert any value to string safely."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ''
    s = str(val).strip()
    return '' if s in ('nan', 'None', 'NaT', 'NaN') else s


def safe_upper(val):
    return safe_str(val).upper()


def col_to_strlist(series):
    """Convert a pandas Series to a plain Python list of strings."""
    return [safe_str(v) for v in series.tolist()]


# ──────────────────────────────────────────────────────────────
# COLUMN FINDER
# ──────────────────────────────────────────────────────────────
def find_col(df, candidates):
    cols_lower = {c.lower().strip(): c for c in df.columns}
    for cand in candidates:
        cl = cand.lower().strip()
        if cl in cols_lower:
            return cols_lower[cl]
    for cand in candidates:
        cl = cand.lower().strip()
        for k, v in cols_lower.items():
            if cl in k:
                return v
    return None


# ──────────────────────────────────────────────────────────────
# DATA PROCESSING
# ──────────────────────────────────────────────────────────────
def process_data(df):
    carrier_col = find_col(df, ['Carrier Name'])
    bol_col = find_col(df, ['Bill of Lading'])
    tracked_col = find_col(df, ['Tracked'])

    if carrier_col is None and bol_col is None:
        st.error("Could not find 'Carrier Name' or 'Bill of Lading' columns.")
        return None, None

    # Filter empty rows using plain python
    keep = []
    for i in range(len(df)):
        has_data = False
        if carrier_col and safe_str(df.iloc[i].get(carrier_col, '')):
            has_data = True
        if bol_col and safe_str(df.iloc[i].get(bol_col, '')):
            has_data = True
        if tracked_col and safe_str(df.iloc[i].get(tracked_col, '')):
            has_data = True
        keep.append(has_data)
    df = df[keep].reset_index(drop=True)

    def gcol(candidates):
        c = find_col(df, candidates)
        if c is None:
            return [''] * len(df)
        return [safe_str(v) for v in df[c].tolist()]

    pickup_name = gcol(['Pickup Name'])
    pickup_cs = gcol(['Pickup City State'])
    pickup_country = gcol(['Pickup Country'])
    dest_name = gcol(['Final Destination Name'])
    dest_cs = gcol(['Final Destination City State'])
    dest_country = gcol(['Final Destination Country'])

    n = len(df)
    pickup_loc = [','.join(filter(None, [pickup_name[i], pickup_cs[i], pickup_country[i]])) for i in range(n)]
    dest_loc = [','.join(filter(None, [dest_name[i], dest_cs[i], dest_country[i]])) for i in range(n)]

    m1_raw = gcol(['Pickup Arrival Milestone (UTC)', 'Pickup Arrival Milestone'])
    m2_raw = gcol(['Pickup Departure Milestone (UTC)', 'Pickup Departure Milestone'])
    m3_raw = gcol(['Final Destination Arrival Milestone (UTC)', 'Final Destination Arrival Milestone'])
    m4_raw = gcol(['Final Destination Departure Milestone (UTC)', 'Final Destination Departure Milestone'])
    tracked_raw = gcol(['Tracked'])
    tracking_error_raw = gcol(['Tracking Error'])
    carrier_raw = gcol(['Carrier Name'])
    bol_raw = gcol(['Bill of Lading'])
    order_raw = gcol(['Order Number'])
    conn_raw = gcol(['Connection Type'])
    method_raw = gcol(['Tracking Method'])
    aeq_raw = gcol(['Active Equipment ID'])
    heq_raw = gcol(['Historical Equipment ID'])
    paw_raw = gcol(['Pickup Appointement Window (UTC)', 'Pickup Appointment Window'])
    daw_raw = gcol(['Delivery Appointement Window (UTC)', 'Delivery Appointment Window'])
    sc_raw = gcol(['Shipment Created (UTC)', 'Shipment Created'])
    tws_raw = gcol(['Tracking Window Start (UTC)', 'Tracking Window Start'])
    twe_raw = gcol(['Tracking Window End (UTC)', 'Tracking Window End'])
    ms_recv = gcol(['# Of Milestones received / # Of Milestones expected'])
    upd_recv = gcol(['# Updates Received'])
    upd10 = gcol(['# Updates Received < 10 mins'])
    nb_exp = gcol(['Nb Intervals Expected'])
    nb_obs = gcol(['Nb Intervals Observed'])
    fsr_raw = gcol(['Final Status Reason'])
    me1 = gcol(['Milestone Error 1'])
    me2 = gcol(['Milestone Error 2'])
    me3 = gcol(['Milestone Error 3'])

    # Build query_df
    query_records = []
    for i in range(n):
        query_records.append({
            'Carrier Name': carrier_raw[i],
            'Bill of Lading': bol_raw[i],
            'Order Number': order_raw[i],
            'Tracked': tracked_raw[i],
            'Connection Type': conn_raw[i],
            'Tracking Method': method_raw[i],
            'Active Equipment ID': aeq_raw[i],
            'Historical Equipment ID': heq_raw[i],
            'Pickup Name': pickup_name[i],
            'Pickup Location': pickup_loc[i],
            'Pickup City State': pickup_cs[i],
            'Pickup Country': pickup_country[i],
            'Pickup Appointement Window (UTC)': paw_raw[i],
            'Final Destination Name': dest_name[i],
            'Final Destination City State': dest_cs[i],
            'Final Destination Country': dest_country[i],
            'Delivery Appointement Window (UTC)': daw_raw[i],
            'Shipment Created (UTC)': sc_raw[i],
            'Tracking Window Start (UTC)': tws_raw[i],
            'Tracking Window End (UTC)': twe_raw[i],
            'Pickup Arrival Milestone (UTC)': m1_raw[i],
            'Pickup Departure Milestone (UTC)': m2_raw[i],
            'Final Destination Arrival Milestone (UTC)': m3_raw[i],
            'Final Destination Departure Milestone (UTC)': m4_raw[i],
            '# Of Milestones received / # Of Milestones expected': ms_recv[i],
            '# Updates Received': upd_recv[i],
            '# Updates Received < 10 mins': upd10[i],
            'Nb Intervals Expected': nb_exp[i],
            'Nb Intervals Observed': nb_obs[i],
            'Final Status Reason': fsr_raw[i],
            'Tracking Error': tracking_error_raw[i],
            'Milestone Error 1': me1[i],
            'Milestone Error 2': me2[i],
            'Milestone Error 3': me3[i],
        })
    query_df = pd.DataFrame(query_records)

    # Build analysis
    def has_ms(val):
        return val not in ('', '0', 'UNKNOWN')

    analysis_records = []
    for i in range(n):
        tracked = tracked_raw[i].upper() == 'TRUE'
        m1 = has_ms(m1_raw[i])
        m2 = has_ms(m2_raw[i])
        m3 = has_ms(m3_raw[i])
        m4 = has_ms(m4_raw[i])

        achieved, missed = [], []
        for flag, lbl in [(m1, 'm1'), (m2, 'm2'), (m3, 'm3'), (m4, 'm4')]:
            (achieved if flag else missed).append(lbl)

        cnt = len(achieved)
        if tracked and cnt == 4:
            ts, mm, ana, p44 = 'Full Tracked', 'Fully Tracked', '', 'Full Tracked'
        elif tracked and cnt > 0:
            ts, mm, ana, p44 = 'Partial Tracked', ', '.join(missed), '', 'Partial Tracked'
        elif tracked:
            ts, mm, ana, p44 = 'Tracked with 0 milestones', 'm1, m2, m3, m4', '', 'Tracked with 0 milestones'
        else:
            err = tracking_error_raw[i]
            ts = 'Tracked with 0 milestones'
            mm = 'm1, m2, m3, m4'
            ana = err
            p44 = err if err else 'Tracked with 0 milestones'

        lane = f"{pickup_cs[i]} -> {dest_cs[i]}" if pickup_cs[i] and dest_cs[i] else ''

        analysis_records.append({
            'Carrier Name': carrier_raw[i],
            'Bill of Lading': bol_raw[i],
            'Order Number': order_raw[i],
            'Tracked': tracked_raw[i],
            'Connection Type': conn_raw[i],
            'Tracking Method': method_raw[i],
            'Active Equipment ID': aeq_raw[i],
            'Historical Equipment ID': heq_raw[i],
            'Lanes': lane,
            'Pickup Location': pickup_loc[i],
            'Destination Location': dest_loc[i],
            'Pickup Appointement Window (UTC)': paw_raw[i],
            'Delivery Appointement Window (UTC)': daw_raw[i],
            'Shipment Created (UTC)': sc_raw[i],
            'Tracking Window Start (UTC)': tws_raw[i],
            'Tracking Window End (UTC)': twe_raw[i],
            'Pickup Arrival Milestone (UTC)': m1_raw[i],
            'Pickup Departure Milestone (UTC)': m2_raw[i],
            'Final Destination Arrival Milestone (UTC)': m3_raw[i],
            'Final Destination Departure Milestone (UTC)': m4_raw[i],
            'Final Status Reason': fsr_raw[i],
            'Tracked Status': ts,
            'Milstone Completeness': '4/4',
            'Milestone Achieved': ', '.join(achieved) if achieved else '',
            'Milestone Missed': mm,
            'Analysis': ana,
            'p44 Analysis': p44,
            'Tracking Error': tracking_error_raw[i],
        })

    return query_df, pd.DataFrame(analysis_records)


# ──────────────────────────────────────────────────────────────
# EXCEL EXPORT (safe for all Python versions)
# ──────────────────────────────────────────────────────────────
def to_excel(sheets_dict):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for name, df in sheets_dict.items():
            sn = name[:31]
            df.to_excel(writer, sheet_name=sn, index=False)
            ws = writer.sheets[sn]
            for i, col in enumerate(df.columns):
                try:
                    max_data = max(len(safe_str(v)) for v in df.iloc[:, i].tolist()) if len(df) > 0 else 0
                except Exception:
                    max_data = 0
                ws.set_column(i, i, min(max(max_data, len(str(col))) + 2, 45))
    return output.getvalue()


# ──────────────────────────────────────────────────────────────
# PIVOT BUILDERS (all use plain python, no .str accessor)
# ──────────────────────────────────────────────────────────────
def _count(series):
    """Value counts as dict, plain python."""
    counts = {}
    for v in series.tolist():
        s = safe_str(v)
        counts[s] = counts.get(s, 0) + 1
    return counts


def _group_count(df, cols):
    """Group by multiple cols, return list of dicts."""
    groups = {}
    for _, row in df.iterrows():
        key = tuple(safe_str(row[c]) for c in cols)
        groups[key] = groups.get(key, 0) + 1
    result = []
    for key, cnt in sorted(groups.items(), key=lambda x: -x[1]):
        d = {c: k for c, k in zip(cols, key)}
        d['Shipments'] = cnt
        result.append(d)
    return result


def pivot_tracking_summary(df):
    total = len(df)
    counts = _count(df['Tracked'].apply(safe_upper))
    rows = [{'Tracked': k, 'Shipments': v, 'Shipment %': f"{v/total*100:.1f}%"} for k, v in sorted(counts.items(), key=lambda x: -x[1])]
    rows.append({'Tracked': 'Grand Total', 'Shipments': total, 'Shipment %': '100%'})
    return pd.DataFrame(rows)


def pivot_tracking_carrier(df):
    total = len(df)
    tmp = df.copy()
    tmp['_tracked'] = [safe_upper(v) for v in tmp['Tracked'].tolist()]
    groups = _group_count(tmp, ['_tracked', 'Carrier Name'])
    for g in groups:
        g['Tracked'] = g.pop('_tracked')
        g['Shipment %'] = f"{g['Shipments']/total*100:.1f}%"
    return pd.DataFrame(groups)


def pivot_milestone_summary(df):
    total = len(df)
    order = ['Full Tracked', 'Partial Tracked', 'Tracked with 0 milestones']
    counts = _count(df['Tracked Status'])
    rows = [{'Tracked Status': s, 'Shipments': counts.get(s, 0), 'Shipment %': f"{counts.get(s,0)/total*100:.1f}%"} for s in order]
    rows.append({'Tracked Status': 'Grand Total', 'Shipments': total, 'Shipment %': '100%'})
    return pd.DataFrame(rows)


def pivot_rca(df):
    total = len(df)
    counts = _count(df['p44 Analysis'])
    rows = [{'P44 Analysis': k, 'Shipments': v, 'Shipment %': f"{v/total*100:.1f}%"} for k, v in sorted(counts.items(), key=lambda x: -x[1])]
    rows.append({'P44 Analysis': 'Grand Total', 'Shipments': total, 'Shipment %': '100%'})
    return pd.DataFrame(rows)


def pivot_ms_carrier(df):
    total = len(df)
    groups = _group_count(df, ['Carrier Name', 'Milestone Missed'])
    for g in groups:
        g['Shipment %'] = f"{g['Shipments']/total*100:.1f}%"
    return pd.DataFrame(groups)


def pivot_lane(df):
    total = len(df)
    groups = _group_count(df, ['Lanes', 'Carrier Name', 'Milestone Missed'])
    for g in groups:
        g['Shipment %'] = f"{g['Shipments']/total*100:.1f}%"
    return pd.DataFrame(groups)


# ──────────────────────────────────────────────────────────────
# BADGE HELPERS
# ──────────────────────────────────────────────────────────────
def badge(text, color='green'):
    return f'<span class="bg bg-{color}">{text}</span>'

def tracked_badge(val):
    return badge(val, 'green') if safe_upper(val) == 'TRUE' else badge(val, 'red')

def status_badge(val):
    if val == 'Full Tracked': return badge(val, 'green')
    if val == 'Partial Tracked': return badge(val, 'orange')
    return badge(val, 'red')

def missed_badge(val):
    return badge(val, 'green') if val == 'Fully Tracked' else badge(val, 'orange')

def p44_badge(val):
    if val == 'Full Tracked': return badge(val, 'green')
    if val == 'Partial Tracked': return badge(val, 'orange')
    if 'milestones' in safe_str(val).lower(): return badge(val, 'blue')
    return badge(val, 'red')

def pct_bar(pct, color='#00875a'):
    return f'<span class="pct-bar" style="width:{max(pct,2):.0f}px;background:{color}"></span>{pct:.1f}%'


# ──────────────────────────────────────────────────────────────
# HTML TABLE RENDERER
# ──────────────────────────────────────────────────────────────
def html_table(df, max_h=400):
    h = f'<div style="max-height:{max_h}px;overflow:auto"><table style="width:100%;border-collapse:collapse;font-size:13px">'
    h += '<thead><tr>'
    for c in df.columns:
        h += f'<th style="background:#f8f9fb;padding:10px 14px;text-align:left;font-weight:600;font-size:12px;text-transform:uppercase;letter-spacing:0.5px;color:#6b778c;border-bottom:2px solid #dfe1e6;white-space:nowrap;position:sticky;top:0;z-index:1">{c}</th>'
    h += '</tr></thead><tbody>'
    for _, row in df.iterrows():
        is_total = 'Total' in safe_str(row.iloc[0])
        style = 'background:#f0f2f5;font-weight:700;border-top:2px solid #dfe1e6' if is_total else ''
        h += f'<tr style="{style}">'
        for v in row:
            h += f'<td style="padding:9px 14px;border-bottom:1px solid #f0f0f0;white-space:nowrap">{safe_str(v)}</td>'
        h += '</tr>'
    h += '</tbody></table></div>'
    return h


# ──────────────────────────────────────────────────────────────
# FILTER HELPER (plain python, no .str accessor)
# ──────────────────────────────────────────────────────────────
def filter_df(df, col, val):
    """Filter dataframe where col equals val, using plain python."""
    mask = [safe_upper(v) == val.upper() for v in df[col].tolist()]
    return df[mask].reset_index(drop=True)


def filter_df_exact(df, col, val):
    """Filter dataframe where col exactly equals val."""
    mask = [safe_str(v) == val for v in df[col].tolist()]
    return df[mask].reset_index(drop=True)


# ──────────────────────────────────────────────────────────────
# MAIN APP
# ──────────────────────────────────────────────────────────────
def main():
    st.markdown("""
    <div class="main-header">
        <div><h1>📊 Data Quality Report</h1><div class="subtitle">project44 Visibility Platform</div></div>
        <div><span class="hdr-badge">Powered by p44</span></div>
    </div>
    """, unsafe_allow_html=True)

    if 'processed' not in st.session_state:
        st.session_state.processed = False

    uploaded = st.file_uploader("Upload your weekly data export (.xlsx)", type=['xlsx', 'xls', 'csv'])

    if uploaded and not st.session_state.processed:
        with st.spinner("Processing data..."):
            try:
                if uploaded.name.endswith('.csv'):
                    raw = pd.read_csv(uploaded)
                else:
                    xls = pd.ExcelFile(uploaded)
                    sheet = 'Data' if 'Data' in xls.sheet_names else xls.sheet_names[0]
                    raw = pd.read_excel(uploaded, sheet_name=sheet)
                q, a = process_data(raw)
                if q is not None:
                    st.session_state.query_df = q
                    st.session_state.analysis_df = a
                    st.session_state.processed = True
                    st.rerun()
            except Exception as e:
                st.error(f"Error: {e}")
                st.code(traceback.format_exc())
                return

    if not st.session_state.processed:
        st.info("👆 Upload your data file to generate the report")
        return

    query_df = st.session_state.query_df
    analysis_df = st.session_state.analysis_df

    # Reset
    _, rc = st.columns([8, 2])
    with rc:
        if st.button("↺ Upload New File", use_container_width=True):
            for k in ['processed', 'query_df', 'analysis_df']:
                st.session_state.pop(k, None)
            st.rerun()

    # ── KPIs ────────────────────────────────────────────────
    total = len(analysis_df)
    tracked_list = [safe_upper(v) for v in analysis_df['Tracked'].tolist()]
    status_list = [safe_str(v) for v in analysis_df['Tracked Status'].tolist()]
    tracked_true = sum(1 for v in tracked_list if v == 'TRUE')
    full = sum(1 for v in status_list if v == 'Full Tracked')
    partial = sum(1 for v in status_list if v == 'Partial Tracked')
    zero_ms = sum(1 for v in status_list if v == 'Tracked with 0 milestones')
    t_pct = (tracked_true / total * 100) if total else 0
    f_pct = (full / total * 100) if total else 0

    st.markdown(f"""
    <div class="kpi-row">
        <div class="kpi-card kpi-blue"><div class="value">{total:,}</div><div class="label">Total Shipments</div></div>
        <div class="kpi-card kpi-green"><div class="value">{t_pct:.1f}%</div><div class="label">Tracked (TRUE)</div></div>
        <div class="kpi-card kpi-green"><div class="value">{full:,}</div><div class="label">Full Tracked</div></div>
        <div class="kpi-card kpi-orange"><div class="value">{partial:,}</div><div class="label">Partial Tracked</div></div>
        <div class="kpi-card kpi-red"><div class="value">{zero_ms:,}</div><div class="label">0 Milestones</div></div>
        <div class="kpi-card"><div class="value">{f_pct:.1f}%</div><div class="label">Full Track Rate</div></div>
    </div>
    """, unsafe_allow_html=True)

    # ── Export Buttons ──────────────────────────────────────
    c1, c2, c3, _ = st.columns([2, 2, 2, 4])
    with c1:
        st.download_button("⬇ Export All to Excel",
            to_excel({'Query': query_df, 'Data Analysis': analysis_df,
                      'DQ-Tracking Summary': pivot_tracking_summary(analysis_df),
                      'Tracking by Carrier': pivot_tracking_carrier(analysis_df),
                      'DQ-Milestone Summary': pivot_milestone_summary(analysis_df),
                      'P44 RCA': pivot_rca(analysis_df),
                      'Lane Analysis': pivot_lane(analysis_df),
                      'Milestone by Carrier': pivot_ms_carrier(analysis_df)}),
            "DQ-Report-Full.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with c2:
        st.download_button("⬇ Tracking Report",
            to_excel({'DQ-Tracking Summary': pivot_tracking_summary(analysis_df),
                      'Tracking by Carrier': pivot_tracking_carrier(analysis_df),
                      'Tracking Detail': analysis_df[['Carrier Name','Bill of Lading','Tracked','Connection Type','Tracking Method','Lanes','Pickup Location','Destination Location','Final Status Reason','Tracked Status','Tracking Error']]}),
            "DQ-Tracking-Report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with c3:
        st.download_button("⬇ Milestone Report",
            to_excel({'DQ-Milestone Summary': pivot_milestone_summary(analysis_df),
                      'P44 RCA': pivot_rca(analysis_df),
                      'Milestone by Carrier': pivot_ms_carrier(analysis_df),
                      'Lane Analysis': pivot_lane(analysis_df),
                      'Milestone Detail': analysis_df}),
            "DQ-Milestone-Report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # ── Carrier list for filters ────────────────────────────
    carriers = sorted(set(safe_str(v) for v in analysis_df['Carrier Name'].tolist() if safe_str(v)))
    carriers_opts = ['All'] + carriers

    # ── TABS ────────────────────────────────────────────────
    tab1, tab2, tab3 = st.tabs(["📡 Tracking Data", "🎯 Milestone Data", "📋 Processed Data"])

    # ═══════════════════════════════════════════════════════
    # TAB 1: TRACKING
    # ═══════════════════════════════════════════════════════
    with tab1:
        col1, col2 = st.columns(2)
        with col1:
            st.markdown('<div class="sec-card"><h3>DQ-Tracking Summary</h3>', unsafe_allow_html=True)
            ts = pivot_tracking_summary(analysis_df)
            ts_d = ts.copy()
            color_map = {'TRUE': '#00875a', 'FALSE': '#de350b'}
            display_tracked = []
            display_pct = []
            for _, r in ts.iterrows():
                t_val = safe_str(r['Tracked'])
                if t_val == 'Grand Total':
                    display_tracked.append(f'<b>{t_val}</b>')
                    display_pct.append('<b>100%</b>')
                else:
                    display_tracked.append(tracked_badge(t_val))
                    display_pct.append(pct_bar(r['Shipments']/total*100, color_map.get(t_val, '#666')))
            ts_d['Tracked'] = display_tracked
            ts_d['Shipment %'] = display_pct
            st.markdown(html_table(ts_d, 250), unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

        with col2:
            st.markdown('<div class="sec-card"><h3>Tracking by Carrier</h3>', unsafe_allow_html=True)
            tc = pivot_tracking_carrier(analysis_df)
            tc_d = tc.copy()
            tc_d['Tracked'] = [tracked_badge(v) for v in tc['Tracked'].tolist()]
            st.markdown(html_table(tc_d, 300), unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

        st.markdown("#### Tracking Detail")
        fc1, fc2, _ = st.columns([2, 3, 5])
        with fc1:
            tf = st.selectbox("Tracked", ["All", "TRUE", "FALSE"], key="t_f")
        with fc2:
            cf = st.selectbox("Carrier", carriers_opts, key="t_c")

        det = analysis_df.copy()
        if tf != "All":
            det = filter_df(det, 'Tracked', tf)
        if cf != "All":
            det = filter_df_exact(det, 'Carrier Name', cf)

        st.caption(f"{len(det)} shipments")
        cols = ['Carrier Name','Bill of Lading','Tracked','Connection Type','Tracking Method','Pickup Location','Destination Location','Lanes','Final Status Reason','Tracked Status','Tracking Error']
        dd = det[cols].copy()
        dd['Tracked'] = [tracked_badge(v) for v in dd['Tracked'].tolist()]
        dd['Tracked Status'] = [status_badge(v) for v in dd['Tracked Status'].tolist()]
        st.markdown(html_table(dd, 500), unsafe_allow_html=True)

    # ═══════════════════════════════════════════════════════
    # TAB 2: MILESTONE
    # ═══════════════════════════════════════════════════════
    with tab2:
        col1, col2 = st.columns(2)
        with col1:
            st.markdown('<div class="sec-card"><h3>DQ-Milestone Summary</h3>', unsafe_allow_html=True)
            ms = pivot_milestone_summary(analysis_df)
            ms_d = ms.copy()
            cm2 = {'Full Tracked':'#00875a','Partial Tracked':'#ff991f','Tracked with 0 milestones':'#de350b'}
            d_status, d_pct = [], []
            for _, r in ms.iterrows():
                sv = safe_str(r['Tracked Status'])
                if sv == 'Grand Total':
                    d_status.append(f'<b>{sv}</b>')
                    d_pct.append('<b>100%</b>')
                else:
                    d_status.append(status_badge(sv))
                    d_pct.append(pct_bar(r['Shipments']/total*100, cm2.get(sv, '#666')))
            ms_d['Tracked Status'] = d_status
            ms_d['Shipment %'] = d_pct
            st.markdown(html_table(ms_d, 250), unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

        with col2:
            st.markdown('<div class="sec-card"><h3>P44 Root Cause Analysis</h3>', unsafe_allow_html=True)
            rca = pivot_rca(analysis_df)
            rca_d = rca.copy()
            rca_d['P44 Analysis'] = [p44_badge(v) if safe_str(v) != 'Grand Total' else f'<b>{v}</b>' for v in rca['P44 Analysis'].tolist()]
            st.markdown(html_table(rca_d, 300), unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

        col3, col4 = st.columns(2)
        with col3:
            st.markdown('<div class="sec-card"><h3>Milestone by Carrier</h3>', unsafe_allow_html=True)
            mc = pivot_ms_carrier(analysis_df)
            mc_d = mc.copy()
            mc_d['Milestone Missed'] = [missed_badge(v) for v in mc['Milestone Missed'].tolist()]
            st.markdown(html_table(mc_d, 300), unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

        with col4:
            st.markdown('<div class="sec-card"><h3>Lane Analysis</h3>', unsafe_allow_html=True)
            la = pivot_lane(analysis_df)
            la_d = la.copy()
            la_d['Milestone Missed'] = [missed_badge(v) for v in la['Milestone Missed'].tolist()]
            st.markdown(html_table(la_d, 300), unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

        st.markdown("#### Milestone Detail")
        mc1, mc2, _ = st.columns([2, 3, 5])
        with mc1:
            sf = st.selectbox("Tracked Status", ['All','Full Tracked','Partial Tracked','Tracked with 0 milestones'], key="m_s")
        with mc2:
            mcf = st.selectbox("Carrier", carriers_opts, key="m_c")

        mdet = analysis_df.copy()
        if sf != "All":
            mdet = filter_df_exact(mdet, 'Tracked Status', sf)
        if mcf != "All":
            mdet = filter_df_exact(mdet, 'Carrier Name', mcf)

        st.caption(f"{len(mdet)} shipments")
        mcols = ['Carrier Name','Bill of Lading','Tracked','Lanes','Pickup Location','Destination Location',
                 'Pickup Arrival Milestone (UTC)','Pickup Departure Milestone (UTC)',
                 'Final Destination Arrival Milestone (UTC)','Final Destination Departure Milestone (UTC)',
                 'Final Status Reason','Tracked Status','Milestone Achieved','Milestone Missed','p44 Analysis']
        md = mdet[mcols].copy()
        md['Tracked'] = [tracked_badge(v) for v in md['Tracked'].tolist()]
        md['Tracked Status'] = [status_badge(v) for v in md['Tracked Status'].tolist()]
        md['Milestone Missed'] = [missed_badge(v) for v in md['Milestone Missed'].tolist()]
        md['p44 Analysis'] = [p44_badge(v) for v in md['p44 Analysis'].tolist()]
        st.markdown(html_table(md, 500), unsafe_allow_html=True)

    # ═══════════════════════════════════════════════════════
    # TAB 3: PROCESSED DATA
    # ═══════════════════════════════════════════════════════
    with tab3:
        st.markdown("#### Full Processed Data (Query Sheet)")
        rc1, rc2, _ = st.columns([2, 3, 5])
        with rc1:
            rtf = st.selectbox("Tracked", ["All", "TRUE", "FALSE"], key="r_t")
        with rc2:
            rcf = st.selectbox("Carrier", carriers_opts, key="r_c")

        rdet = query_df.copy()
        if rtf != "All":
            rdet = filter_df(rdet, 'Tracked', rtf)
        if rcf != "All":
            rdet = filter_df_exact(rdet, 'Carrier Name', rcf)

        st.caption(f"{len(rdet)} shipments")
        rd = rdet.copy()
        rd['Tracked'] = [tracked_badge(v) for v in rd['Tracked'].tolist()]
        st.markdown(html_table(rd, 600), unsafe_allow_html=True)


if __name__ == '__main__':
    main()

import streamlit as st
import pandas as pd
import io

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
# CUSTOM CSS — matches the HTML version exactly
# ──────────────────────────────────────────────────────────────
st.markdown("""
<style>
/* ── Header bar ─────────────────────────────────────────── */
.main-header {
    background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%);
    color: white;
    padding: 16px 32px;
    border-radius: 10px;
    margin-bottom: 20px;
    display: flex;
    justify-content: space-between;
    align-items: center;
}
.main-header h1 { font-size: 22px; font-weight: 600; margin: 0; letter-spacing: 0.5px; color: white; }
.main-header .subtitle { font-size: 12px; opacity: 0.7; margin-top: 2px; }
.main-header .hdr-badge {
    background: rgba(255,255,255,0.12);
    padding: 6px 14px;
    border-radius: 20px;
    font-size: 13px;
}

/* ── KPI cards row ──────────────────────────────────────── */
.kpi-row { display: flex; gap: 12px; flex-wrap: wrap; margin-bottom: 18px; }
.kpi-card {
    flex: 1; min-width: 130px;
    background: white;
    padding: 16px 20px;
    border-radius: 10px;
    text-align: center;
    box-shadow: 0 1px 3px rgba(0,0,0,0.08);
    border: 1px solid #eee;
}
.kpi-card .value { font-size: 28px; font-weight: 700; }
.kpi-card .label {
    font-size: 11px; color: #6b778c;
    text-transform: uppercase; letter-spacing: 0.8px; margin-top: 4px;
}
.kpi-green .value  { color: #00875a; }
.kpi-red .value    { color: #de350b; }
.kpi-orange .value { color: #ff991f; }
.kpi-blue .value   { color: #0046FF; }

/* ── Status badges (inline in tables) ───────────────────── */
.bg { display: inline-block; padding: 3px 10px; border-radius: 12px; font-size: 12px; font-weight: 600; }
.bg-green  { background: #e3fcef; color: #006644; }
.bg-red    { background: #ffebe6; color: #bf2600; }
.bg-orange { background: #fff7e6; color: #974f0c; }
.bg-blue   { background: #deebff; color: #0747a6; }

/* ── Section cards ──────────────────────────────────────── */
.sec-card {
    background: white;
    border-radius: 10px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.08);
    padding: 20px;
    margin-bottom: 16px;
    border: 1px solid #eee;
}
.sec-card h3 { font-size: 15px; font-weight: 600; margin: 0 0 12px 0; color: #172b4d; }

/* ── Percentage bar inside pivot tables ─────────────────── */
.pct-bar {
    display: inline-block; height: 8px; border-radius: 4px; min-width: 4px; vertical-align: middle; margin-right: 6px;
}

/* ── Total row bold ─────────────────────────────────────── */
.total-row td { font-weight: 700 !important; background: #f0f2f5 !important; }

/* ── Streamlit overrides ────────────────────────────────── */
.block-container { padding-top: 1rem; max-width: 1400px; }
div[data-testid="stTabs"] button { font-weight: 500; }
div[data-testid="stMetric"] {
    background: white; padding: 16px; border-radius: 10px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.08);
}
div[data-testid="stMetric"] label {
    font-size: 11px !important; text-transform: uppercase; letter-spacing: 0.8px;
}
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
header {visibility: hidden;}
</style>
""", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────
# HELPER: flexible column finder
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
# DATA PROCESSING — same Power Query logic
# ──────────────────────────────────────────────────────────────
def process_data(df):
    carrier_col = find_col(df, ['Carrier Name'])
    bol_col = find_col(df, ['Bill of Lading'])
    tracked_col = find_col(df, ['Tracked'])

    if carrier_col is None and bol_col is None:
        st.error("Could not find 'Carrier Name' or 'Bill of Lading' columns.")
        return None, None

    # Filter empty rows
    mask = pd.Series([False] * len(df), index=df.index)
    if carrier_col:
        mask = mask | (df[carrier_col].notna() & (df[carrier_col].astype(str).str.strip() != ''))
    if bol_col:
        mask = mask | (df[bol_col].notna() & (df[bol_col].astype(str).str.strip() != ''))
    if tracked_col:
        mask = mask | (df[tracked_col].notna() & (df[tracked_col].astype(str).str.strip() != ''))
    df = df[mask].reset_index(drop=True)

    def gcol(candidates):
        c = find_col(df, candidates)
        return df[c].astype(str).replace('nan', '').replace('None', '') if c else pd.Series([''] * len(df))

    pickup_name = gcol(['Pickup Name'])
    pickup_cs = gcol(['Pickup City State'])
    pickup_country = gcol(['Pickup Country'])
    dest_name = gcol(['Final Destination Name'])
    dest_cs = gcol(['Final Destination City State'])
    dest_country = gcol(['Final Destination Country'])

    query_df = pd.DataFrame({
        'Carrier Name': gcol(['Carrier Name']),
        'Bill of Lading': gcol(['Bill of Lading']),
        'Order Number': gcol(['Order Number']),
        'Tracked': gcol(['Tracked']),
        'Connection Type': gcol(['Connection Type']),
        'Tracking Method': gcol(['Tracking Method']),
        'Active Equipment ID': gcol(['Active Equipment ID']),
        'Historical Equipment ID': gcol(['Historical Equipment ID']),
        'Pickup Name': pickup_name,
        'Pickup Location': pickup_name + ',' + pickup_cs + ',' + pickup_country,
        'Pickup City State': pickup_cs,
        'Pickup Country': pickup_country,
        'Pickup Appointement Window (UTC)': gcol(['Pickup Appointement Window (UTC)', 'Pickup Appointment Window']),
        'Final Destination Name': dest_name,
        'Final Destination City State': dest_cs,
        'Final Destination Country': dest_country,
        'Delivery Appointement Window (UTC)': gcol(['Delivery Appointement Window (UTC)', 'Delivery Appointment Window']),
        'Shipment Created (UTC)': gcol(['Shipment Created (UTC)', 'Shipment Created']),
        'Tracking Window Start (UTC)': gcol(['Tracking Window Start (UTC)', 'Tracking Window Start']),
        'Tracking Window End (UTC)': gcol(['Tracking Window End (UTC)', 'Tracking Window End']),
        'Pickup Arrival Milestone (UTC)': gcol(['Pickup Arrival Milestone (UTC)', 'Pickup Arrival Milestone']),
        'Pickup Departure Milestone (UTC)': gcol(['Pickup Departure Milestone (UTC)', 'Pickup Departure Milestone']),
        'Final Destination Arrival Milestone (UTC)': gcol(['Final Destination Arrival Milestone (UTC)', 'Final Destination Arrival']),
        'Final Destination Departure Milestone (UTC)': gcol(['Final Destination Departure Milestone (UTC)', 'Final Destination Departure']),
        '# Of Milestones received / # Of Milestones expected': gcol(['# Of Milestones received / # Of Milestones expected']),
        '# Updates Received': gcol(['# Updates Received']),
        '# Updates Received < 10 mins': gcol(['# Updates Received < 10 mins']),
        'Nb Intervals Expected': gcol(['Nb Intervals Expected']),
        'Nb Intervals Observed': gcol(['Nb Intervals Observed']),
        'Final Status Reason': gcol(['Final Status Reason']),
        'Tracking Error': gcol(['Tracking Error']),
        'Milestone Error 1': gcol(['Milestone Error 1']),
        'Milestone Error 2': gcol(['Milestone Error 2']),
        'Milestone Error 3': gcol(['Milestone Error 3']),
    })

    # Build Analysis
    def has_ms(val):
        s = str(val).strip()
        return s not in ('', '0', 'UNKNOWN', 'None', 'nan', 'NaT', 'NaN')

    records = []
    for idx, row in query_df.iterrows():
        tracked = str(row['Tracked']).strip().upper() == 'TRUE'
        m1 = has_ms(row['Pickup Arrival Milestone (UTC)'])
        m2 = has_ms(row['Pickup Departure Milestone (UTC)'])
        m3 = has_ms(row['Final Destination Arrival Milestone (UTC)'])
        m4 = has_ms(row['Final Destination Departure Milestone (UTC)'])

        achieved, missed = [], []
        for flag, label in [(m1, 'm1'), (m2, 'm2'), (m3, 'm3'), (m4, 'm4')]:
            (achieved if flag else missed).append(label)

        n = len(achieved)
        if tracked and n == 4:
            ts, mm, ana, p44 = 'Full Tracked', 'Fully Tracked', '', 'Full Tracked'
        elif tracked and n > 0:
            ts, mm, ana, p44 = 'Partial Tracked', ', '.join(missed), '', 'Partial Tracked'
        elif tracked:
            ts, mm, ana, p44 = 'Tracked with 0 milestones', 'm1, m2, m3, m4', '', 'Tracked with 0 milestones'
        else:
            err = str(row['Tracking Error']).strip()
            err = '' if err in ('', 'nan', 'None') else err
            ts = 'Tracked with 0 milestones'
            mm = 'm1, m2, m3, m4'
            ana = err
            p44 = err if err else 'Tracked with 0 milestones'

        pcs = str(row['Pickup City State']).strip()
        dcs = str(row['Final Destination City State']).strip()
        lane = f"{pcs} -> {dcs}" if pcs and dcs else ''
        dest_loc = ','.join(filter(None, [str(row['Final Destination Name']).strip(), dcs, str(row['Final Destination Country']).strip()]))

        records.append({
            'Carrier Name': row['Carrier Name'],
            'Bill of Lading': row['Bill of Lading'],
            'Order Number': row['Order Number'],
            'Tracked': row['Tracked'],
            'Connection Type': row['Connection Type'],
            'Tracking Method': row['Tracking Method'],
            'Active Equipment ID': row['Active Equipment ID'],
            'Historical Equipment ID': row['Historical Equipment ID'],
            'Lanes': lane,
            'Pickup Location': row['Pickup Location'],
            'Destination Location': dest_loc,
            'Pickup Appointement Window (UTC)': row['Pickup Appointement Window (UTC)'],
            'Delivery Appointement Window (UTC)': row['Delivery Appointement Window (UTC)'],
            'Shipment Created (UTC)': row['Shipment Created (UTC)'],
            'Tracking Window Start (UTC)': row['Tracking Window Start (UTC)'],
            'Tracking Window End (UTC)': row['Tracking Window End (UTC)'],
            'Pickup Arrival Milestone (UTC)': row['Pickup Arrival Milestone (UTC)'],
            'Pickup Departure Milestone (UTC)': row['Pickup Departure Milestone (UTC)'],
            'Final Destination Arrival Milestone (UTC)': row['Final Destination Arrival Milestone (UTC)'],
            'Final Destination Departure Milestone (UTC)': row['Final Destination Departure Milestone (UTC)'],
            'Final Status Reason': row['Final Status Reason'],
            'Tracked Status': ts,
            'Milstone Completeness': '4/4',
            'Milestone Achieved': ', '.join(achieved) if achieved else '',
            'Milestone Missed': mm,
            'Analysis': ana,
            'p44 Analysis': p44,
            'Tracking Error': row['Tracking Error'],
        })

    return query_df, pd.DataFrame(records)


# ──────────────────────────────────────────────────────────────
# EXCEL EXPORT (fixed for all pandas/python versions)
# ──────────────────────────────────────────────────────────────
def to_excel(sheets_dict):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for name, df in sheets_dict.items():
            sheet_name = name[:31]
            df.to_excel(writer, sheet_name=sheet_name, index=False)
            ws = writer.sheets[sheet_name]
            for i, col in enumerate(df.columns):
                try:
                    lengths = df[col].astype(str).str.len()
                    max_data = int(lengths.max()) if len(lengths) > 0 and not lengths.isna().all() else 0
                except Exception:
                    max_data = 0
                ws.set_column(i, i, min(max(max_data, len(str(col))) + 2, 45))
    return output.getvalue()


# ──────────────────────────────────────────────────────────────
# PIVOT TABLE BUILDERS
# ──────────────────────────────────────────────────────────────
def pivot_tracking_summary(df):
    total = len(df)
    g = df['Tracked'].str.upper().value_counts().reset_index()
    g.columns = ['Tracked', 'Shipments']
    g['Shipment %'] = (g['Shipments'] / total * 100).round(1).astype(str) + '%'
    g = pd.concat([g, pd.DataFrame([{'Tracked': 'Grand Total', 'Shipments': total, 'Shipment %': '100%'}])], ignore_index=True)
    return g

def pivot_tracking_carrier(df):
    total = len(df)
    g = df.groupby([df['Tracked'].str.upper(), 'Carrier Name']).size().reset_index(name='Shipments')
    g.columns = ['Tracked', 'Carrier Name', 'Shipments']
    g = g.sort_values('Shipments', ascending=False)
    g['Shipment %'] = (g['Shipments'] / total * 100).round(1).astype(str) + '%'
    return g

def pivot_milestone_summary(df):
    total = len(df)
    order = ['Full Tracked', 'Partial Tracked', 'Tracked with 0 milestones']
    g = df['Tracked Status'].value_counts().reindex(order, fill_value=0).reset_index()
    g.columns = ['Tracked Status', 'Shipments']
    g['Shipment %'] = (g['Shipments'] / total * 100).round(1).astype(str) + '%'
    g = pd.concat([g, pd.DataFrame([{'Tracked Status': 'Grand Total', 'Shipments': total, 'Shipment %': '100%'}])], ignore_index=True)
    return g

def pivot_rca(df):
    total = len(df)
    g = df['p44 Analysis'].value_counts().reset_index()
    g.columns = ['P44 Analysis', 'Shipments']
    g['Shipment %'] = (g['Shipments'] / total * 100).round(1).astype(str) + '%'
    g = pd.concat([g, pd.DataFrame([{'P44 Analysis': 'Grand Total', 'Shipments': total, 'Shipment %': '100%'}])], ignore_index=True)
    return g

def pivot_ms_carrier(df):
    total = len(df)
    g = df.groupby(['Carrier Name', 'Milestone Missed']).size().reset_index(name='Shipments')
    g = g.sort_values('Shipments', ascending=False)
    g['Shipment %'] = (g['Shipments'] / total * 100).round(1).astype(str) + '%'
    return g

def pivot_lane(df):
    total = len(df)
    g = df.groupby(['Lanes', 'Carrier Name', 'Milestone Missed']).size().reset_index(name='Shipments')
    g = g.sort_values('Shipments', ascending=False)
    g['Shipment %'] = (g['Shipments'] / total * 100).round(1).astype(str) + '%'
    return g


# ──────────────────────────────────────────────────────────────
# HTML TABLE RENDERER (styled like the HTML page)
# ──────────────────────────────────────────────────────────────
def badge(text, color='green'):
    return f'<span class="bg bg-{color}">{text}</span>'

def tracked_badge(val):
    return badge(val, 'green') if str(val).upper() == 'TRUE' else badge(val, 'red')

def status_badge(val):
    if val == 'Full Tracked': return badge(val, 'green')
    if val == 'Partial Tracked': return badge(val, 'orange')
    return badge(val, 'red')

def missed_badge(val):
    return badge(val, 'green') if val == 'Fully Tracked' else badge(val, 'orange')

def p44_badge(val):
    if val == 'Full Tracked': return badge(val, 'green')
    if val == 'Partial Tracked': return badge(val, 'orange')
    if 'milestones' in str(val).lower(): return badge(val, 'blue')
    return badge(val, 'red')

def pct_bar(pct, color='#00875a'):
    return f'<span class="pct-bar" style="width:{max(pct, 2)}px;background:{color}"></span>{pct:.1f}%'

def render_html_table(df, max_height=400):
    """Render a pandas DataFrame as a styled HTML table."""
    html = f'<div style="max-height:{max_height}px;overflow:auto"><table style="width:100%;border-collapse:collapse;font-size:13px">'
    html += '<thead><tr>'
    for col in df.columns:
        html += f'<th style="background:#f8f9fb;padding:10px 14px;text-align:left;font-weight:600;font-size:12px;text-transform:uppercase;letter-spacing:0.5px;color:#6b778c;border-bottom:2px solid #dfe1e6;white-space:nowrap;position:sticky;top:0;z-index:1">{col}</th>'
    html += '</tr></thead><tbody>'
    for _, row in df.iterrows():
        is_total = 'Grand Total' in str(row.iloc[0]) or 'Total' in str(row.iloc[0])
        bg = 'background:#f0f2f5;font-weight:700;border-top:2px solid #dfe1e6' if is_total else ''
        html += f'<tr style="{bg}">'
        for val in row:
            html += f'<td style="padding:9px 14px;border-bottom:1px solid #f0f0f0;white-space:nowrap">{val}</td>'
        html += '</tr>'
    html += '</tbody></table></div>'
    return html


# ──────────────────────────────────────────────────────────────
# MAIN APP
# ──────────────────────────────────────────────────────────────
def main():
    # Header
    st.markdown("""
    <div class="main-header">
        <div>
            <h1>📊 Data Quality Report</h1>
            <div class="subtitle">project44 Visibility Platform</div>
        </div>
        <div><span class="hdr-badge">Powered by p44</span></div>
    </div>
    """, unsafe_allow_html=True)

    # Session state
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
                    cust_col = find_col(raw, ['Customer Tenant Name'])
                    st.session_state.customer = str(raw[cust_col].dropna().iloc[0]) if cust_col and len(raw[cust_col].dropna()) > 0 else ''
                    st.rerun()
            except Exception as e:
                st.error(f"Error processing file: {e}")
                import traceback
                st.code(traceback.format_exc())
                return

    if not st.session_state.processed:
        st.info("👆 Upload your data file to generate the report")
        return

    query_df = st.session_state.query_df
    analysis_df = st.session_state.analysis_df

    # Reset button
    _, rc = st.columns([8, 2])
    with rc:
        if st.button("↺ Upload New File", use_container_width=True):
            for k in ['processed', 'query_df', 'analysis_df', 'customer']:
                st.session_state.pop(k, None)
            st.rerun()

    # ── KPI Strip ───────────────────────────────────────────
    total = len(analysis_df)
    tracked_true = len(analysis_df[analysis_df['Tracked'].str.upper() == 'TRUE'])
    full = len(analysis_df[analysis_df['Tracked Status'] == 'Full Tracked'])
    partial = len(analysis_df[analysis_df['Tracked Status'] == 'Partial Tracked'])
    zero_ms = len(analysis_df[analysis_df['Tracked Status'] == 'Tracked with 0 milestones'])
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

    # ── Export All ──────────────────────────────────────────
    all_excel = to_excel({
        'Query': query_df,
        'Data Analysis': analysis_df,
        'DQ-Tracking Summary': pivot_tracking_summary(analysis_df),
        'Tracking by Carrier': pivot_tracking_carrier(analysis_df),
        'DQ-Milestone Summary': pivot_milestone_summary(analysis_df),
        'P44 RCA': pivot_rca(analysis_df),
        'Lane Analysis': pivot_lane(analysis_df),
        'Milestone by Carrier': pivot_ms_carrier(analysis_df),
    })
    c1, c2, c3, _ = st.columns([2, 2, 2, 4])
    with c1:
        st.download_button("⬇ Export All to Excel", all_excel, "DQ-Report-Full.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with c2:
        tr_excel = to_excel({
            'DQ-Tracking Summary': pivot_tracking_summary(analysis_df),
            'Tracking by Carrier': pivot_tracking_carrier(analysis_df),
            'Tracking Detail': analysis_df[['Carrier Name','Bill of Lading','Tracked','Connection Type','Tracking Method','Lanes','Pickup Location','Destination Location','Final Status Reason','Tracked Status','Tracking Error']],
        })
        st.download_button("⬇ Tracking Report", tr_excel, "DQ-Tracking-Report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with c3:
        ms_excel = to_excel({
            'DQ-Milestone Summary': pivot_milestone_summary(analysis_df),
            'P44 RCA': pivot_rca(analysis_df),
            'Milestone by Carrier': pivot_ms_carrier(analysis_df),
            'Lane Analysis': pivot_lane(analysis_df),
            'Milestone Detail': analysis_df,
        })
        st.download_button("⬇ Milestone Report", ms_excel, "DQ-Milestone-Report.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

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
            ts_display = ts.copy()
            ts_display['Tracked'] = ts_display['Tracked'].apply(lambda x: tracked_badge(x) if x not in ('Grand Total',) else f'<b>{x}</b>')
            color_map = {'TRUE': '#00875a', 'FALSE': '#de350b'}
            ts_display['Shipment %'] = ts.apply(lambda r: pct_bar(r['Shipments']/total*100, color_map.get(r['Tracked'], '#666')) if r['Tracked'] != 'Grand Total' else '<b>100%</b>', axis=1)
            st.markdown(render_html_table(ts_display, 250), unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

        with col2:
            st.markdown('<div class="sec-card"><h3>Tracking by Carrier</h3>', unsafe_allow_html=True)
            tc = pivot_tracking_carrier(analysis_df)
            tc_display = tc.copy()
            tc_display['Tracked'] = tc_display['Tracked'].apply(tracked_badge)
            st.markdown(render_html_table(tc_display, 300), unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

        # Detail with filters
        st.markdown("#### Tracking Detail")
        fc1, fc2, _ = st.columns([2, 3, 5])
        with fc1:
            tf = st.selectbox("Tracked", ["All", "TRUE", "FALSE"], key="t_f")
        with fc2:
            carriers = ['All'] + sorted(analysis_df['Carrier Name'].unique().tolist())
            cf = st.selectbox("Carrier", carriers, key="t_c")

        det = analysis_df.copy()
        if tf != "All":
            det = det[det['Tracked'].str.upper() == tf]
        if cf != "All":
            det = det[det['Carrier Name'] == cf]

        st.caption(f"{len(det)} shipments")
        cols = ['Carrier Name', 'Bill of Lading', 'Tracked', 'Connection Type', 'Tracking Method',
                'Pickup Location', 'Destination Location', 'Lanes', 'Final Status Reason', 'Tracked Status', 'Tracking Error']
        dd = det[cols].copy()
        dd['Tracked'] = dd['Tracked'].apply(tracked_badge)
        dd['Tracked Status'] = dd['Tracked Status'].apply(status_badge)
        st.markdown(render_html_table(dd, 500), unsafe_allow_html=True)

    # ═══════════════════════════════════════════════════════
    # TAB 2: MILESTONE
    # ═══════════════════════════════════════════════════════
    with tab2:
        col1, col2 = st.columns(2)
        with col1:
            st.markdown('<div class="sec-card"><h3>DQ-Milestone Summary</h3>', unsafe_allow_html=True)
            ms = pivot_milestone_summary(analysis_df)
            ms_d = ms.copy()
            ms_d['Tracked Status'] = ms_d['Tracked Status'].apply(lambda x: status_badge(x) if x != 'Grand Total' else f'<b>{x}</b>')
            color_map2 = {'Full Tracked': '#00875a', 'Partial Tracked': '#ff991f', 'Tracked with 0 milestones': '#de350b'}
            ms_d['Shipment %'] = ms.apply(lambda r: pct_bar(r['Shipments']/total*100, color_map2.get(r['Tracked Status'], '#666')) if r['Tracked Status'] != 'Grand Total' else '<b>100%</b>', axis=1)
            st.markdown(render_html_table(ms_d, 250), unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

        with col2:
            st.markdown('<div class="sec-card"><h3>P44 Root Cause Analysis</h3>', unsafe_allow_html=True)
            rca = pivot_rca(analysis_df)
            rca_d = rca.copy()
            rca_d['P44 Analysis'] = rca_d['P44 Analysis'].apply(lambda x: p44_badge(x) if x != 'Grand Total' else f'<b>{x}</b>')
            st.markdown(render_html_table(rca_d, 300), unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

        col3, col4 = st.columns(2)
        with col3:
            st.markdown('<div class="sec-card"><h3>Milestone by Carrier</h3>', unsafe_allow_html=True)
            mc = pivot_ms_carrier(analysis_df)
            mc_d = mc.copy()
            mc_d['Milestone Missed'] = mc_d['Milestone Missed'].apply(missed_badge)
            st.markdown(render_html_table(mc_d, 300), unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

        with col4:
            st.markdown('<div class="sec-card"><h3>Lane Analysis</h3>', unsafe_allow_html=True)
            la = pivot_lane(analysis_df)
            la_d = la.copy()
            la_d['Milestone Missed'] = la_d['Milestone Missed'].apply(missed_badge)
            st.markdown(render_html_table(la_d, 300), unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

        # Milestone Detail with filters
        st.markdown("#### Milestone Detail")
        mc1, mc2, _ = st.columns([2, 3, 5])
        with mc1:
            sf = st.selectbox("Tracked Status", ['All', 'Full Tracked', 'Partial Tracked', 'Tracked with 0 milestones'], key="m_s")
        with mc2:
            mcf = st.selectbox("Carrier", carriers, key="m_c")

        mdet = analysis_df.copy()
        if sf != "All":
            mdet = mdet[mdet['Tracked Status'] == sf]
        if mcf != "All":
            mdet = mdet[mdet['Carrier Name'] == mcf]

        st.caption(f"{len(mdet)} shipments")
        mcols = ['Carrier Name', 'Bill of Lading', 'Tracked', 'Lanes', 'Pickup Location', 'Destination Location',
                 'Pickup Arrival Milestone (UTC)', 'Pickup Departure Milestone (UTC)',
                 'Final Destination Arrival Milestone (UTC)', 'Final Destination Departure Milestone (UTC)',
                 'Final Status Reason', 'Tracked Status', 'Milestone Achieved', 'Milestone Missed', 'p44 Analysis']
        md = mdet[mcols].copy()
        md['Tracked'] = md['Tracked'].apply(tracked_badge)
        md['Tracked Status'] = md['Tracked Status'].apply(status_badge)
        md['Milestone Missed'] = md['Milestone Missed'].apply(missed_badge)
        md['p44 Analysis'] = md['p44 Analysis'].apply(p44_badge)
        st.markdown(render_html_table(md, 500), unsafe_allow_html=True)

    # ═══════════════════════════════════════════════════════
    # TAB 3: PROCESSED DATA
    # ═══════════════════════════════════════════════════════
    with tab3:
        st.markdown("#### Full Processed Data (Query Sheet)")
        rc1, rc2, _ = st.columns([2, 3, 5])
        with rc1:
            rtf = st.selectbox("Tracked", ["All", "TRUE", "FALSE"], key="r_t")
        with rc2:
            rcf = st.selectbox("Carrier", carriers, key="r_c")

        rdet = query_df.copy()
        if rtf != "All":
            rdet = rdet[rdet['Tracked'].str.upper() == rtf]
        if rcf != "All":
            rdet = rdet[rdet['Carrier Name'] == rcf]

        st.caption(f"{len(rdet)} shipments")
        rd = rdet.copy()
        rd['Tracked'] = rd['Tracked'].apply(tracked_badge)
        st.markdown(render_html_table(rd, 600), unsafe_allow_html=True)


if __name__ == '__main__':
    main()

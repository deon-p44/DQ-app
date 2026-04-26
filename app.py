import streamlit as st
import pandas as pd
import io
from collections import Counter

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
# CUSTOM CSS — matches the HTML version's dark header + card theme
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
.main-header .badge {
    background: rgba(255,255,255,0.12);
    padding: 6px 14px;
    border-radius: 20px;
    font-size: 13px;
}

/* ── KPI cards ──────────────────────────────────────────── */
.kpi-row { display: flex; gap: 12px; flex-wrap: wrap; margin-bottom: 18px; }
.kpi-card {
    flex: 1; min-width: 130px;
    background: white;
    padding: 16px 20px;
    border-radius: 10px;
    text-align: center;
    box-shadow: 0 1px 3px rgba(0,0,0,0.08);
}
.kpi-card .value { font-size: 28px; font-weight: 700; }
.kpi-card .label { font-size: 11px; color: #6b778c; text-transform: uppercase; letter-spacing: 0.8px; margin-top: 4px; }
.kpi-green .value { color: #00875a; }
.kpi-red .value   { color: #de350b; }
.kpi-orange .value{ color: #ff991f; }
.kpi-blue .value  { color: #0046FF; }

/* ── Status badges ──────────────────────────────────────── */
.badge-green  { background: #e3fcef; color: #006644; padding: 3px 10px; border-radius: 12px; font-size: 12px; font-weight: 600; }
.badge-red    { background: #ffebe6; color: #bf2600; padding: 3px 10px; border-radius: 12px; font-size: 12px; font-weight: 600; }
.badge-orange { background: #fff7e6; color: #974f0c; padding: 3px 10px; border-radius: 12px; font-size: 12px; font-weight: 600; }
.badge-blue   { background: #deebff; color: #0747a6; padding: 3px 10px; border-radius: 12px; font-size: 12px; font-weight: 600; }

/* ── Section cards ──────────────────────────────────────── */
.section-card {
    background: white;
    border-radius: 10px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.08);
    padding: 20px;
    margin-bottom: 16px;
}
.section-card h3 { font-size: 15px; font-weight: 600; margin-bottom: 12px; color: #172b4d; }

/* ── Table styling ──────────────────────────────────────── */
.dataframe { font-size: 13px !important; }
div[data-testid="stDataFrame"] { border-radius: 8px; overflow: hidden; }

/* ── Streamlit overrides ────────────────────────────────── */
.block-container { padding-top: 1rem; }
div[data-testid="stTabs"] button { font-weight: 500; }
div[data-testid="stMetric"] { background: white; padding: 16px; border-radius: 10px; box-shadow: 0 1px 3px rgba(0,0,0,0.08); }
div[data-testid="stMetric"] label { font-size: 11px !important; text-transform: uppercase; letter-spacing: 0.8px; }

/* Hide Streamlit footer and menu */
#MainMenu {visibility: hidden;}
footer {visibility: hidden;}
</style>
""", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────
# DATA PROCESSING — same Power Query logic as HTML version
# ──────────────────────────────────────────────────────────────
def find_col(df, candidates):
    """Flexibly find column name from list of candidates."""
    cols_lower = {c.lower().strip(): c for c in df.columns}
    for cand in candidates:
        cl = cand.lower().strip()
        if cl in cols_lower:
            return cols_lower[cl]
    # Partial match
    for cand in candidates:
        cl = cand.lower().strip()
        for k, v in cols_lower.items():
            if cl in k:
                return v
    return None


def process_data(df):
    """Apply Power Query transformation logic."""

    # ── Step 1: Filter out empty rows ───────────────────────
    carrier_col = find_col(df, ['Carrier Name'])
    bol_col = find_col(df, ['Bill of Lading'])
    tracked_col = find_col(df, ['Tracked'])

    if carrier_col is None and bol_col is None:
        st.error("Could not find 'Carrier Name' or 'Bill of Lading' columns. Please check your file.")
        return None, None

    mask = pd.Series([False] * len(df))
    if carrier_col:
        mask = mask | df[carrier_col].notna().astype(bool) & (df[carrier_col] != '')
    if bol_col:
        mask = mask | df[bol_col].notna().astype(bool) & (df[bol_col] != '')
    if tracked_col:
        mask = mask | df[tracked_col].notna().astype(bool) & (df[tracked_col] != '')

    df = df[mask].reset_index(drop=True)

    # ── Step 2: Build Query dataframe ───────────────────────
    def gcol(candidates):
        c = find_col(df, candidates)
        return df[c] if c else pd.Series([''] * len(df))

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
        'Pickup Location': pickup_name.astype(str).str.cat([pickup_cs.astype(str), pickup_country.astype(str)], sep=','),
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
        'Final Destination Arrival Milestone (UTC)': gcol(['Final Destination Arrival Milestone (UTC)', 'Final Destination Arrival Milestone']),
        'Final Destination Departure Milestone (UTC)': gcol(['Final Destination Departure Milestone (UTC)', 'Final Destination Departure Milestone']),
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

    # ── Step 3: Build Analysis dataframe ────────────────────
    def has_milestone(val):
        if pd.isna(val):
            return False
        s = str(val).strip()
        return s not in ('', '0', 'UNKNOWN', 'None', 'nan', 'NaT')

    records = []
    for idx, row in query_df.iterrows():
        tracked = str(row['Tracked']).strip().upper() == 'TRUE'

        m1 = has_milestone(row['Pickup Arrival Milestone (UTC)'])
        m2 = has_milestone(row['Pickup Departure Milestone (UTC)'])
        m3 = has_milestone(row['Final Destination Arrival Milestone (UTC)'])
        m4 = has_milestone(row['Final Destination Departure Milestone (UTC)'])

        achieved = []
        missed = []
        for flag, label in [(m1, 'm1'), (m2, 'm2'), (m3, 'm3'), (m4, 'm4')]:
            (achieved if flag else missed).append(label)

        total_achieved = len(achieved)

        if tracked and total_achieved == 4:
            tracked_status = 'Full Tracked'
            milestone_missed = 'Fully Tracked'
            analysis = ''
            p44_analysis = 'Full Tracked'
        elif tracked and total_achieved > 0:
            tracked_status = 'Partial Tracked'
            milestone_missed = ', '.join(missed)
            analysis = ''
            p44_analysis = 'Partial Tracked'
        elif tracked and total_achieved == 0:
            tracked_status = 'Tracked with 0 milestones'
            milestone_missed = 'm1, m2, m3, m4'
            analysis = ''
            p44_analysis = 'Tracked with 0 milestones'
        else:
            tracked_status = 'Tracked with 0 milestones'
            milestone_missed = 'm1, m2, m3, m4'
            err = str(row['Tracking Error']) if pd.notna(row['Tracking Error']) else ''
            analysis = err
            p44_analysis = err if err else 'Tracked with 0 milestones'

        pcs = str(row['Pickup City State']) if pd.notna(row['Pickup City State']) else ''
        dcs = str(row['Final Destination City State']) if pd.notna(row['Final Destination City State']) else ''
        lane = f"{pcs} -> {dcs}" if pcs and dcs else ''

        dn = str(row['Final Destination Name']) if pd.notna(row['Final Destination Name']) else ''
        dc = str(row['Final Destination Country']) if pd.notna(row['Final Destination Country']) else ''
        dest_loc = ','.join(filter(None, [dn, dcs, dc]))

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
            'Tracked Status': tracked_status,
            'Milstone Completeness': '4/4',
            'Milestone Achieved': ', '.join(achieved) if achieved else '',
            'Milestone Missed': milestone_missed,
            'Analysis': analysis,
            'p44 Analysis': p44_analysis,
        })

    analysis_df = pd.DataFrame(records)
    return query_df, analysis_df


def to_excel_download(sheets_dict, filename):
    """Create downloadable Excel with multiple sheets."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        for name, df in sheets_dict.items():
            df.to_excel(writer, sheet_name=name[:31], index=False)
            worksheet = writer.sheets[name[:31]]
            for i, col in enumerate(df.columns):
                max_len = max(df[col].astype(str).map(len).max(), len(col)) + 2
                worksheet.set_column(i, i, min(max_len, 40))
    return output.getvalue()


def badge_html(text, color='green'):
    return f'<span class="badge-{color}">{text}</span>'


# ──────────────────────────────────────────────────────────────
# MAIN APP
# ──────────────────────────────────────────────────────────────
def main():
    # ── Header ──────────────────────────────────────────────
    st.markdown("""
    <div class="main-header">
        <div>
            <h1>📊 Data Quality Report</h1>
            <div class="subtitle">project44 Visibility Platform</div>
        </div>
        <div>
            <span class="badge">Powered by p44</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # ── File Upload ─────────────────────────────────────────
    if 'processed' not in st.session_state:
        st.session_state.processed = False

    uploaded_file = st.file_uploader(
        "Upload your weekly data export (.xlsx)",
        type=['xlsx', 'xls', 'csv'],
        help="Upload the raw data export from the p44 platform. The system will automatically process and generate reports."
    )

    if uploaded_file is not None and not st.session_state.processed:
        with st.spinner("Processing data..."):
            try:
                if uploaded_file.name.endswith('.csv'):
                    raw_df = pd.read_csv(uploaded_file)
                else:
                    # Try to read "Data" sheet first, fall back to first sheet
                    xls = pd.ExcelFile(uploaded_file)
                    sheet = 'Data' if 'Data' in xls.sheet_names else xls.sheet_names[0]
                    raw_df = pd.read_excel(uploaded_file, sheet_name=sheet)

                query_df, analysis_df = process_data(raw_df)
                if query_df is not None:
                    st.session_state.query_df = query_df
                    st.session_state.analysis_df = analysis_df
                    st.session_state.processed = True
                    # Detect customer name
                    cust_col = find_col(raw_df, ['Customer Tenant Name', 'Tenant Name'])
                    if cust_col:
                        names = raw_df[cust_col].dropna().unique()
                        st.session_state.customer = names[0] if len(names) > 0 else ''
                    else:
                        st.session_state.customer = ''
                    st.rerun()
            except Exception as e:
                st.error(f"Error processing file: {e}")
                return

    if not st.session_state.processed:
        st.info("👆 Upload your data file to generate the report")
        return

    # ── Data loaded — show report ───────────────────────────
    query_df = st.session_state.query_df
    analysis_df = st.session_state.analysis_df
    customer = st.session_state.get('customer', '')

    # Reset button
    col_r1, col_r2 = st.columns([8, 2])
    with col_r2:
        if st.button("↺ Upload New File", type="secondary", use_container_width=True):
            st.session_state.processed = False
            st.session_state.pop('query_df', None)
            st.session_state.pop('analysis_df', None)
            st.rerun()

    # ── KPI Strip ───────────────────────────────────────────
    total = len(analysis_df)
    tracked_true = len(analysis_df[analysis_df['Tracked'].astype(str).str.upper() == 'TRUE'])
    full_tracked = len(analysis_df[analysis_df['Tracked Status'] == 'Full Tracked'])
    partial = len(analysis_df[analysis_df['Tracked Status'] == 'Partial Tracked'])
    zero_ms = len(analysis_df[analysis_df['Tracked Status'] == 'Tracked with 0 milestones'])
    tracking_pct = (tracked_true / total * 100) if total else 0
    full_pct = (full_tracked / total * 100) if total else 0

    k1, k2, k3, k4, k5, k6 = st.columns(6)
    k1.metric("Total Shipments", f"{total:,}")
    k2.metric("Tracked (TRUE)", f"{tracking_pct:.1f}%")
    k3.metric("Full Tracked", f"{full_tracked:,}")
    k4.metric("Partial Tracked", f"{partial:,}")
    k5.metric("0 Milestones", f"{zero_ms:,}")
    k6.metric("Full Track Rate", f"{full_pct:.1f}%")

    st.divider()

    # ── Tabs ────────────────────────────────────────────────
    tab1, tab2, tab3 = st.tabs(["📡 Tracking Data", "🎯 Milestone Data", "📋 Processed Data"])

    # ════════════════════════════════════════════════════════
    # TAB 1: TRACKING DATA
    # ════════════════════════════════════════════════════════
    with tab1:
        # Export button
        tracking_excel = to_excel_download({
            'DQ-Tracking Summary': build_tracking_summary(analysis_df),
            'Tracking by Carrier': build_tracking_carrier(analysis_df),
            'Tracking Detail': analysis_df[['Carrier Name','Bill of Lading','Order Number','Tracked','Connection Type','Tracking Method','Lanes','Pickup Location','Destination Location','Final Status Reason','Tracked Status']].copy(),
        }, 'tracking')
        st.download_button("⬇ Export Tracking Report", tracking_excel, "DQ-Tracking-Report.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        col1, col2 = st.columns(2)

        with col1:
            st.markdown('<div class="section-card"><h3>DQ-Tracking Summary</h3>', unsafe_allow_html=True)
            summary = build_tracking_summary(analysis_df)
            st.dataframe(summary, use_container_width=True, hide_index=True)
            st.markdown('</div>', unsafe_allow_html=True)

        with col2:
            st.markdown('<div class="section-card"><h3>Tracking by Carrier</h3>', unsafe_allow_html=True)
            carrier_summary = build_tracking_carrier(analysis_df)
            st.dataframe(carrier_summary, use_container_width=True, hide_index=True, height=300)
            st.markdown('</div>', unsafe_allow_html=True)

        # Detail table with filters
        st.markdown("### Tracking Detail")
        fc1, fc2, fc3 = st.columns([2, 3, 2])
        with fc1:
            tracked_filter = st.selectbox("Tracked", ["All", "TRUE", "FALSE"], key="t_tracked")
        with fc2:
            carriers = ['All'] + sorted(analysis_df['Carrier Name'].dropna().unique().tolist())
            carrier_filter = st.selectbox("Carrier", carriers, key="t_carrier")

        detail = analysis_df.copy()
        if tracked_filter != "All":
            detail = detail[detail['Tracked'].astype(str).str.upper() == tracked_filter]
        if carrier_filter != "All":
            detail = detail[detail['Carrier Name'] == carrier_filter]

        tracking_cols = ['Carrier Name', 'Bill of Lading', 'Tracked', 'Connection Type', 'Tracking Method',
                         'Pickup Location', 'Destination Location', 'Lanes', 'Final Status Reason', 'Tracked Status']
        # Add tracking error from query_df
        detail_display = detail[tracking_cols].copy()
        st.caption(f"{len(detail_display)} shipments")
        st.dataframe(detail_display, use_container_width=True, hide_index=True, height=400)

    # ════════════════════════════════════════════════════════
    # TAB 2: MILESTONE DATA
    # ════════════════════════════════════════════════════════
    with tab2:
        milestone_excel = to_excel_download({
            'DQ-Milestone Summary': build_milestone_summary(analysis_df),
            'P44 RCA': build_rca(analysis_df),
            'Milestone by Carrier': build_milestone_carrier(analysis_df),
            'Lane Analysis': build_lane_analysis(analysis_df),
            'Milestone Detail': analysis_df.copy(),
        }, 'milestone')
        st.download_button("⬇ Export Milestone Report", milestone_excel, "DQ-Milestone-Report.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        col1, col2 = st.columns(2)

        with col1:
            st.markdown('<div class="section-card"><h3>DQ-Milestone Summary</h3>', unsafe_allow_html=True)
            ms_summary = build_milestone_summary(analysis_df)
            st.dataframe(ms_summary, use_container_width=True, hide_index=True)
            st.markdown('</div>', unsafe_allow_html=True)

        with col2:
            st.markdown('<div class="section-card"><h3>P44 Root Cause Analysis</h3>', unsafe_allow_html=True)
            rca = build_rca(analysis_df)
            st.dataframe(rca, use_container_width=True, hide_index=True, height=300)
            st.markdown('</div>', unsafe_allow_html=True)

        col3, col4 = st.columns(2)

        with col3:
            st.markdown('<div class="section-card"><h3>Milestone by Carrier</h3>', unsafe_allow_html=True)
            ms_carrier = build_milestone_carrier(analysis_df)
            st.dataframe(ms_carrier, use_container_width=True, hide_index=True, height=300)
            st.markdown('</div>', unsafe_allow_html=True)

        with col4:
            st.markdown('<div class="section-card"><h3>Lane Analysis</h3>', unsafe_allow_html=True)
            lane = build_lane_analysis(analysis_df)
            st.dataframe(lane, use_container_width=True, hide_index=True, height=300)
            st.markdown('</div>', unsafe_allow_html=True)

        # Milestone Detail
        st.markdown("### Milestone Detail")
        mc1, mc2, mc3 = st.columns([2, 3, 2])
        with mc1:
            statuses = ['All', 'Full Tracked', 'Partial Tracked', 'Tracked with 0 milestones']
            ms_status_filter = st.selectbox("Tracked Status", statuses, key="m_status")
        with mc2:
            ms_carrier_filter = st.selectbox("Carrier", carriers, key="m_carrier")

        ms_detail = analysis_df.copy()
        if ms_status_filter != "All":
            ms_detail = ms_detail[ms_detail['Tracked Status'] == ms_status_filter]
        if ms_carrier_filter != "All":
            ms_detail = ms_detail[ms_detail['Carrier Name'] == ms_carrier_filter]

        ms_cols = ['Carrier Name', 'Bill of Lading', 'Tracked', 'Lanes', 'Pickup Location', 'Destination Location',
                   'Pickup Arrival Milestone (UTC)', 'Pickup Departure Milestone (UTC)',
                   'Final Destination Arrival Milestone (UTC)', 'Final Destination Departure Milestone (UTC)',
                   'Final Status Reason', 'Tracked Status', 'Milestone Achieved', 'Milestone Missed', 'p44 Analysis']
        st.caption(f"{len(ms_detail)} shipments")
        st.dataframe(ms_detail[ms_cols], use_container_width=True, hide_index=True, height=400)

    # ════════════════════════════════════════════════════════
    # TAB 3: PROCESSED DATA
    # ════════════════════════════════════════════════════════
    with tab3:
        full_excel = to_excel_download({
            'Query': query_df,
            'Data Analysis': analysis_df,
            'DQ-Tracking Summary': build_tracking_summary(analysis_df),
            'Tracking by Carrier': build_tracking_carrier(analysis_df),
            'DQ-Milestone Summary': build_milestone_summary(analysis_df),
            'P44 RCA': build_rca(analysis_df),
            'Lane Analysis': build_lane_analysis(analysis_df),
            'Milestone by Carrier': build_milestone_carrier(analysis_df),
        }, 'full')
        st.download_button("⬇ Export All to Excel", full_excel, "DQ-Report-Full.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

        st.markdown("### Full Processed Data (Query Sheet)")
        rc1, rc2, rc3 = st.columns([2, 3, 2])
        with rc1:
            raw_tracked = st.selectbox("Tracked", ["All", "TRUE", "FALSE"], key="r_tracked")
        with rc2:
            raw_carrier = st.selectbox("Carrier", carriers, key="r_carrier")

        raw_display = query_df.copy()
        if raw_tracked != "All":
            raw_display = raw_display[raw_display['Tracked'].astype(str).str.upper() == raw_tracked]
        if raw_carrier != "All":
            raw_display = raw_display[raw_display['Carrier Name'] == raw_carrier]

        st.caption(f"{len(raw_display)} shipments")
        st.dataframe(raw_display, use_container_width=True, hide_index=True, height=500)


# ──────────────────────────────────────────────────────────────
# PIVOT BUILDERS
# ──────────────────────────────────────────────────────────────
def build_tracking_summary(df):
    total = len(df)
    groups = df['Tracked'].astype(str).str.upper().value_counts().reset_index()
    groups.columns = ['Tracked', 'Shipments']
    groups['Shipment %'] = (groups['Shipments'] / total * 100).round(1).astype(str) + '%'
    grand = pd.DataFrame([{'Tracked': 'Grand Total', 'Shipments': total, 'Shipment %': '100%'}])
    return pd.concat([groups, grand], ignore_index=True)


def build_tracking_carrier(df):
    total = len(df)
    groups = df.groupby([df['Tracked'].astype(str).str.upper(), 'Carrier Name']).size().reset_index(name='Shipments')
    groups.columns = ['Tracked', 'Carrier Name', 'Shipments']
    groups = groups.sort_values('Shipments', ascending=False)
    groups['Shipment %'] = (groups['Shipments'] / total * 100).round(1).astype(str) + '%'
    return groups


def build_milestone_summary(df):
    total = len(df)
    order = ['Full Tracked', 'Partial Tracked', 'Tracked with 0 milestones']
    groups = df['Tracked Status'].value_counts().reindex(order, fill_value=0).reset_index()
    groups.columns = ['Tracked Status', 'Shipments']
    groups['Shipment %'] = (groups['Shipments'] / total * 100).round(1).astype(str) + '%'
    grand = pd.DataFrame([{'Tracked Status': 'Grand Total', 'Shipments': total, 'Shipment %': '100%'}])
    return pd.concat([groups, grand], ignore_index=True)


def build_rca(df):
    total = len(df)
    groups = df['p44 Analysis'].value_counts().reset_index()
    groups.columns = ['P44 Analysis', 'Shipments']
    groups['Shipment %'] = (groups['Shipments'] / total * 100).round(1).astype(str) + '%'
    grand = pd.DataFrame([{'P44 Analysis': 'Grand Total', 'Shipments': total, 'Shipment %': '100%'}])
    return pd.concat([groups, grand], ignore_index=True)


def build_milestone_carrier(df):
    total = len(df)
    groups = df.groupby(['Carrier Name', 'Milestone Missed']).size().reset_index(name='Shipments')
    groups = groups.sort_values('Shipments', ascending=False)
    groups['Shipment %'] = (groups['Shipments'] / total * 100).round(1).astype(str) + '%'
    return groups


def build_lane_analysis(df):
    total = len(df)
    groups = df.groupby(['Lanes', 'Carrier Name', 'Milestone Missed']).size().reset_index(name='Shipments')
    groups = groups.sort_values('Shipments', ascending=False)
    groups['Shipment %'] = (groups['Shipments'] / total * 100).round(1).astype(str) + '%'
    return groups


# ──────────────────────────────────────────────────────────────
if __name__ == '__main__':
    main()

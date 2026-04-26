# ──────────────────────────────────────────────────────────────
# SELF-INSTALL DEPENDENCIES (bulletproof for Streamlit Cloud)
# ──────────────────────────────────────────────────────────────
import subprocess, sys

for pkg in ["openpyxl", "xlsxwriter"]:
    try:
        __import__(pkg)
    except ImportError:
        subprocess.check_call([sys.executable, "-m", "pip", "install", pkg, "-q"])

import pandas as pd
try:
    pd.options.future.infer_string = False
except Exception:
    pass

import streamlit as st
import io
import traceback as tb

# ──────────────────────────────────────────────────────────────
# PAGE CONFIG
# ──────────────────────────────────────────────────────────────
st.set_page_config(page_title="Data Quality Report", page_icon="📊", layout="wide", initial_sidebar_state="collapsed")

# ──────────────────────────────────────────────────────────────
# CSS
# ──────────────────────────────────────────────────────────────
st.markdown("""
<style>
.mh{background:linear-gradient(135deg,#1a1a2e 0%,#16213e 100%);color:#fff;padding:16px 32px;border-radius:10px;margin-bottom:20px;display:flex;justify-content:space-between;align-items:center}
.mh h1{font-size:22px;font-weight:600;margin:0;color:#fff}.mh .sub{font-size:12px;opacity:.7;margin-top:2px}
.mh .hb{background:rgba(255,255,255,.12);padding:6px 14px;border-radius:20px;font-size:13px}
.kr{display:flex;gap:12px;flex-wrap:wrap;margin-bottom:18px}
.kc{flex:1;min-width:130px;background:#fff;padding:16px 20px;border-radius:10px;text-align:center;box-shadow:0 1px 3px rgba(0,0,0,.08);border:1px solid #eee}
.kc .v{font-size:28px;font-weight:700}.kc .l{font-size:11px;color:#6b778c;text-transform:uppercase;letter-spacing:.8px;margin-top:4px}
.kg .v{color:#00875a}.krd .v{color:#de350b}.ko .v{color:#ff991f}.kb .v{color:#0046FF}
.bg{display:inline-block;padding:3px 10px;border-radius:12px;font-size:12px;font-weight:600}
.bg-g{background:#e3fcef;color:#006644}.bg-r{background:#ffebe6;color:#bf2600}
.bg-o{background:#fff7e6;color:#974f0c}.bg-b{background:#deebff;color:#0747a6}
.sc{background:#fff;border-radius:10px;box-shadow:0 1px 3px rgba(0,0,0,.08);padding:20px;margin-bottom:16px;border:1px solid #eee}
.sc h3{font-size:15px;font-weight:600;margin:0 0 12px;color:#172b4d}
.pb{display:inline-block;height:8px;border-radius:4px;min-width:4px;vertical-align:middle;margin-right:6px}
.block-container{padding-top:1rem;max-width:1400px}
div[data-testid="stMetric"]{background:#fff;padding:16px;border-radius:10px;box-shadow:0 1px 3px rgba(0,0,0,.08)}
div[data-testid="stMetric"] label{font-size:11px!important;text-transform:uppercase;letter-spacing:.8px}
#MainMenu{visibility:hidden}footer{visibility:hidden}header{visibility:hidden}
</style>
""", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────
# HELPERS
# ──────────────────────────────────────────────────────────────
def ss(val):
    if val is None:
        return ''
    if isinstance(val, float) and pd.isna(val):
        return ''
    s = str(val).strip()
    return '' if s in ('nan','None','NaT','NaN') else s

def su(val):
    return ss(val).upper()

def find_col(df, cands):
    lo = {c.lower().strip(): c for c in df.columns}
    for c in cands:
        cl = c.lower().strip()
        if cl in lo: return lo[cl]
    for c in cands:
        cl = c.lower().strip()
        for k, v in lo.items():
            if cl in k: return v
    return None


# ──────────────────────────────────────────────────────────────
# PROCESS DATA
# ──────────────────────────────────────────────────────────────
def process(df):
    cc = find_col(df, ['Carrier Name'])
    bc = find_col(df, ['Bill of Lading'])
    tc = find_col(df, ['Tracked'])
    if cc is None and bc is None:
        st.error("Cannot find 'Carrier Name' or 'Bill of Lading' columns.")
        return None, None

    keep = []
    for i in range(len(df)):
        r = df.iloc[i]
        ok = False
        if cc and ss(r.get(cc, '')): ok = True
        if bc and ss(r.get(bc, '')): ok = True
        if tc and ss(r.get(tc, '')): ok = True
        keep.append(ok)
    df = df[keep].reset_index(drop=True)
    n = len(df)

    def gc(cands):
        c = find_col(df, cands)
        return [ss(df.iloc[i][c]) for i in range(n)] if c else ['']*n

    cn = gc(['Carrier Name']); bl = gc(['Bill of Lading']); on = gc(['Order Number'])
    tr = gc(['Tracked']); ct = gc(['Connection Type']); tm = gc(['Tracking Method'])
    ae = gc(['Active Equipment ID']); he = gc(['Historical Equipment ID'])
    pn = gc(['Pickup Name']); pcs = gc(['Pickup City State']); pc = gc(['Pickup Country'])
    dn = gc(['Final Destination Name']); dcs = gc(['Final Destination City State']); dc = gc(['Final Destination Country'])
    paw = gc(['Pickup Appointement Window (UTC)','Pickup Appointment Window'])
    daw = gc(['Delivery Appointement Window (UTC)','Delivery Appointment Window'])
    scr = gc(['Shipment Created (UTC)','Shipment Created'])
    tws = gc(['Tracking Window Start (UTC)','Tracking Window Start'])
    twe = gc(['Tracking Window End (UTC)','Tracking Window End'])
    m1r = gc(['Pickup Arrival Milestone (UTC)','Pickup Arrival Milestone'])
    m2r = gc(['Pickup Departure Milestone (UTC)','Pickup Departure Milestone'])
    m3r = gc(['Final Destination Arrival Milestone (UTC)','Final Destination Arrival'])
    m4r = gc(['Final Destination Departure Milestone (UTC)','Final Destination Departure'])
    mre = gc(['# Of Milestones received / # Of Milestones expected'])
    ur = gc(['# Updates Received']); u10 = gc(['# Updates Received < 10 mins'])
    nie = gc(['Nb Intervals Expected']); nio = gc(['Nb Intervals Observed'])
    fsr = gc(['Final Status Reason']); ter = gc(['Tracking Error'])
    me1 = gc(['Milestone Error 1']); me2 = gc(['Milestone Error 2']); me3 = gc(['Milestone Error 3'])

    def hm(v): return v not in ('','0','UNKNOWN')

    qrows, arows = [], []
    for i in range(n):
        pl = ','.join(filter(None,[pn[i],pcs[i],pc[i]]))
        dl = ','.join(filter(None,[dn[i],dcs[i],dc[i]]))
        qrows.append({
            'Carrier Name':cn[i],'Bill of Lading':bl[i],'Order Number':on[i],'Tracked':tr[i],
            'Connection Type':ct[i],'Tracking Method':tm[i],'Active Equipment ID':ae[i],
            'Historical Equipment ID':he[i],'Pickup Name':pn[i],'Pickup Location':pl,
            'Pickup City State':pcs[i],'Pickup Country':pc[i],
            'Pickup Appointement Window (UTC)':paw[i],'Final Destination Name':dn[i],
            'Final Destination City State':dcs[i],'Final Destination Country':dc[i],
            'Delivery Appointement Window (UTC)':daw[i],'Shipment Created (UTC)':scr[i],
            'Tracking Window Start (UTC)':tws[i],'Tracking Window End (UTC)':twe[i],
            'Pickup Arrival Milestone (UTC)':m1r[i],'Pickup Departure Milestone (UTC)':m2r[i],
            'Final Destination Arrival Milestone (UTC)':m3r[i],
            'Final Destination Departure Milestone (UTC)':m4r[i],
            '# Of Milestones received / # Of Milestones expected':mre[i],
            '# Updates Received':ur[i],'# Updates Received < 10 mins':u10[i],
            'Nb Intervals Expected':nie[i],'Nb Intervals Observed':nio[i],
            'Final Status Reason':fsr[i],'Tracking Error':ter[i],
            'Milestone Error 1':me1[i],'Milestone Error 2':me2[i],'Milestone Error 3':me3[i],
        })

        trkd = tr[i].upper() == 'TRUE'
        a1,a2,a3,a4 = hm(m1r[i]),hm(m2r[i]),hm(m3r[i]),hm(m4r[i])
        ach = [l for f,l in [(a1,'m1'),(a2,'m2'),(a3,'m3'),(a4,'m4')] if f]
        mis = [l for f,l in [(a1,'m1'),(a2,'m2'),(a3,'m3'),(a4,'m4')] if not f]
        nc = len(ach)

        if trkd and nc==4: ts,mm,ana,p44='Full Tracked','Fully Tracked','','Full Tracked'
        elif trkd and nc>0: ts,mm,ana,p44='Partial Tracked',', '.join(mis),'','Partial Tracked'
        elif trkd: ts,mm,ana,p44='Tracked with 0 milestones','m1, m2, m3, m4','','Tracked with 0 milestones'
        else:
            e = ter[i]
            ts='Tracked with 0 milestones'; mm='m1, m2, m3, m4'; ana=e; p44=e if e else 'Tracked with 0 milestones'

        lane = f"{pcs[i]} -> {dcs[i]}" if pcs[i] and dcs[i] else ''
        arows.append({
            'Carrier Name':cn[i],'Bill of Lading':bl[i],'Order Number':on[i],'Tracked':tr[i],
            'Connection Type':ct[i],'Tracking Method':tm[i],'Active Equipment ID':ae[i],
            'Historical Equipment ID':he[i],'Lanes':lane,'Pickup Location':pl,
            'Destination Location':dl,
            'Pickup Appointement Window (UTC)':paw[i],'Delivery Appointement Window (UTC)':daw[i],
            'Shipment Created (UTC)':scr[i],'Tracking Window Start (UTC)':tws[i],
            'Tracking Window End (UTC)':twe[i],
            'Pickup Arrival Milestone (UTC)':m1r[i],'Pickup Departure Milestone (UTC)':m2r[i],
            'Final Destination Arrival Milestone (UTC)':m3r[i],
            'Final Destination Departure Milestone (UTC)':m4r[i],
            'Final Status Reason':fsr[i],'Tracked Status':ts,'Milstone Completeness':'4/4',
            'Milestone Achieved':', '.join(ach) if ach else '','Milestone Missed':mm,
            'Analysis':ana,'p44 Analysis':p44,'Tracking Error':ter[i],
        })

    return pd.DataFrame(qrows), pd.DataFrame(arows)


# ──────────────────────────────────────────────────────────────
# EXCEL EXPORT
# ──────────────────────────────────────────────────────────────
def to_xl(sheets):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='xlsxwriter') as w:
        for nm, df in sheets.items():
            sn = nm[:31]; df.to_excel(w, sheet_name=sn, index=False)
            ws = w.sheets[sn]
            for i, c in enumerate(df.columns):
                try: ml = max(max(len(ss(v)) for v in df.iloc[:,i].tolist()),0) if len(df)>0 else 0
                except: ml = 0
                ws.set_column(i, i, min(max(ml, len(str(c)))+2, 45))
    return buf.getvalue()


# ──────────────────────────────────────────────────────────────
# PIVOTS (all plain python)
# ──────────────────────────────────────────────────────────────
def _cnt(lst):
    d = {}
    for v in lst: d[v] = d.get(v,0)+1
    return d

def _gcnt(df, cols):
    g = {}
    for i in range(len(df)):
        k = tuple(ss(df.iloc[i][c]) for c in cols)
        g[k] = g.get(k,0)+1
    return sorted(g.items(), key=lambda x:-x[1])

def pv_track(df):
    t = len(df); c = _cnt([su(v) for v in df['Tracked'].tolist()])
    r = [{'Tracked':k,'Shipments':v,'Shipment %':f"{v/t*100:.1f}%"} for k,v in sorted(c.items(),key=lambda x:-x[1])]
    r.append({'Tracked':'Grand Total','Shipments':t,'Shipment %':'100%'})
    return pd.DataFrame(r)

def pv_track_carr(df):
    t = len(df)
    tmp = [(su(df.iloc[i]['Tracked']), ss(df.iloc[i]['Carrier Name'])) for i in range(len(df))]
    c = {}
    for tk,cn in tmp: c[(tk,cn)] = c.get((tk,cn),0)+1
    r = [{'Tracked':k[0],'Carrier Name':k[1],'Shipments':v,'Shipment %':f"{v/t*100:.1f}%"} for k,v in sorted(c.items(),key=lambda x:-x[1])]
    return pd.DataFrame(r)

def pv_ms(df):
    t = len(df); c = _cnt([ss(v) for v in df['Tracked Status'].tolist()])
    order = ['Full Tracked','Partial Tracked','Tracked with 0 milestones']
    r = [{'Tracked Status':s,'Shipments':c.get(s,0),'Shipment %':f"{c.get(s,0)/t*100:.1f}%"} for s in order]
    r.append({'Tracked Status':'Grand Total','Shipments':t,'Shipment %':'100%'})
    return pd.DataFrame(r)

def pv_rca(df):
    t = len(df); c = _cnt([ss(v) for v in df['p44 Analysis'].tolist()])
    r = [{'P44 Analysis':k,'Shipments':v,'Shipment %':f"{v/t*100:.1f}%"} for k,v in sorted(c.items(),key=lambda x:-x[1])]
    r.append({'P44 Analysis':'Grand Total','Shipments':t,'Shipment %':'100%'})
    return pd.DataFrame(r)

def pv_mc(df):
    t = len(df); g = _gcnt(df, ['Carrier Name','Milestone Missed'])
    return pd.DataFrame([{'Carrier Name':k[0],'Milestone Missed':k[1],'Shipments':v,'Shipment %':f"{v/t*100:.1f}%"} for k,v in g])

def pv_lane(df):
    t = len(df); g = _gcnt(df, ['Lanes','Carrier Name','Milestone Missed'])
    return pd.DataFrame([{'Lanes':k[0],'Carrier Name':k[1],'Milestone Missed':k[2],'Shipments':v,'Shipment %':f"{v/t*100:.1f}%"} for k,v in g])


# ──────────────────────────────────────────────────────────────
# BADGES & TABLE
# ──────────────────────────────────────────────────────────────
def bg(t,c): return f'<span class="bg bg-{c}">{t}</span>'
def b_tr(v): return bg(v,'g') if su(v)=='TRUE' else bg(v,'r')
def b_st(v):
    if v=='Full Tracked': return bg(v,'g')
    if v=='Partial Tracked': return bg(v,'o')
    return bg(v,'r')
def b_mm(v): return bg(v,'g') if v=='Fully Tracked' else bg(v,'o')
def b_p4(v):
    if v=='Full Tracked': return bg(v,'g')
    if v=='Partial Tracked': return bg(v,'o')
    if 'milestones' in ss(v).lower(): return bg(v,'b')
    return bg(v,'r')
def pb(p,c='#00875a'): return f'<span class="pb" style="width:{max(p,2):.0f}px;background:{c}"></span>{p:.1f}%'

def ht(df, mh=400):
    h = f'<div style="max-height:{mh}px;overflow:auto"><table style="width:100%;border-collapse:collapse;font-size:13px">'
    h += '<thead><tr>'
    for c in df.columns:
        h += f'<th style="background:#f8f9fb;padding:10px 14px;text-align:left;font-weight:600;font-size:12px;text-transform:uppercase;letter-spacing:.5px;color:#6b778c;border-bottom:2px solid #dfe1e6;white-space:nowrap;position:sticky;top:0;z-index:1">{c}</th>'
    h += '</tr></thead><tbody>'
    for i in range(len(df)):
        r = df.iloc[i]
        tot = 'Total' in ss(r.iloc[0])
        s = 'background:#f0f2f5;font-weight:700;border-top:2px solid #dfe1e6' if tot else ''
        h += f'<tr style="{s}">'
        for v in r: h += f'<td style="padding:9px 14px;border-bottom:1px solid #f0f0f0;white-space:nowrap">{ss(v)}</td>'
        h += '</tr>'
    h += '</tbody></table></div>'
    return h


# ──────────────────────────────────────────────────────────────
# FILTER
# ──────────────────────────────────────────────────────────────
def flt(df, col, val, exact=False):
    m = []
    for i in range(len(df)):
        v = ss(df.iloc[i][col])
        m.append(v == val if exact else v.upper() == val.upper())
    return df[m].reset_index(drop=True)


# ──────────────────────────────────────────────────────────────
# MAIN
# ──────────────────────────────────────────────────────────────
def main():
    st.markdown('<div class="mh"><div><h1>📊 Data Quality Report</h1><div class="sub">project44 Visibility Platform</div></div><div><span class="hb">Powered by p44</span></div></div>', unsafe_allow_html=True)

    if 'ok' not in st.session_state: st.session_state.ok = False

    up = st.file_uploader("Upload your weekly data export (.xlsx)", type=['xlsx','xls','csv'])

    if up and not st.session_state.ok:
        with st.spinner("Processing..."):
            try:
                if up.name.endswith('.csv'):
                    raw = pd.read_csv(up)
                else:
                    xls = pd.ExcelFile(up, engine='openpyxl')
                    sh = 'Data' if 'Data' in xls.sheet_names else xls.sheet_names[0]
                    raw = pd.read_excel(xls, sheet_name=sh, engine='openpyxl')
                q, a = process(raw)
                if q is not None:
                    st.session_state.q = q; st.session_state.a = a; st.session_state.ok = True
                    st.rerun()
            except Exception as e:
                st.error(f"Error: {e}")
                st.code(tb.format_exc())
                return

    if not st.session_state.ok:
        st.info("👆 Upload your data file to generate the report")
        return

    q = st.session_state.q; a = st.session_state.a
    _, rc = st.columns([8,2])
    with rc:
        if st.button("↺ Upload New File", use_container_width=True):
            st.session_state.ok = False; st.session_state.pop('q',None); st.session_state.pop('a',None); st.rerun()

    # KPIs
    tot = len(a)
    tl = [su(v) for v in a['Tracked'].tolist()]
    sl = [ss(v) for v in a['Tracked Status'].tolist()]
    tt = sum(1 for v in tl if v=='TRUE')
    ft = sum(1 for v in sl if v=='Full Tracked')
    pt = sum(1 for v in sl if v=='Partial Tracked')
    zm = sum(1 for v in sl if v=='Tracked with 0 milestones')
    tp = tt/tot*100 if tot else 0; fp = ft/tot*100 if tot else 0

    st.markdown(f"""<div class="kr">
    <div class="kc kb"><div class="v">{tot:,}</div><div class="l">Total Shipments</div></div>
    <div class="kc kg"><div class="v">{tp:.1f}%</div><div class="l">Tracked (TRUE)</div></div>
    <div class="kc kg"><div class="v">{ft:,}</div><div class="l">Full Tracked</div></div>
    <div class="kc ko"><div class="v">{pt:,}</div><div class="l">Partial Tracked</div></div>
    <div class="kc krd"><div class="v">{zm:,}</div><div class="l">0 Milestones</div></div>
    <div class="kc"><div class="v">{fp:.1f}%</div><div class="l">Full Track Rate</div></div>
    </div>""", unsafe_allow_html=True)

    # Exports
    c1,c2,c3,_ = st.columns([2,2,2,4])
    with c1:
        st.download_button("⬇ Export All", to_xl({'Query':q,'Data Analysis':a,'DQ-Tracking':pv_track(a),'Tracking Carrier':pv_track_carr(a),'DQ-Milestone':pv_ms(a),'P44 RCA':pv_rca(a),'Lane Analysis':pv_lane(a),'Milestone Carrier':pv_mc(a)}), "DQ-Report-Full.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with c2:
        st.download_button("⬇ Tracking", to_xl({'DQ-Tracking':pv_track(a),'Tracking Carrier':pv_track_carr(a),'Detail':a[['Carrier Name','Bill of Lading','Tracked','Connection Type','Tracking Method','Lanes','Pickup Location','Destination Location','Final Status Reason','Tracked Status','Tracking Error']]}), "DQ-Tracking.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with c3:
        st.download_button("⬇ Milestone", to_xl({'DQ-Milestone':pv_ms(a),'P44 RCA':pv_rca(a),'Milestone Carrier':pv_mc(a),'Lane Analysis':pv_lane(a),'Detail':a}), "DQ-Milestone.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    carrs = ['All'] + sorted(set(ss(v) for v in a['Carrier Name'].tolist() if ss(v)))

    # TABS
    t1,t2,t3 = st.tabs(["📡 Tracking Data","🎯 Milestone Data","📋 Processed Data"])

    # ── TAB 1 ───────────────────────────────────────────────
    with t1:
        c1,c2 = st.columns(2)
        with c1:
            st.markdown('<div class="sc"><h3>DQ-Tracking Summary</h3>', unsafe_allow_html=True)
            ts = pv_track(a); td = ts.copy()
            td['Tracked'] = [b_tr(v) if v!='Grand Total' else f'<b>{v}</b>' for v in ts['Tracked'].tolist()]
            cm = {'TRUE':'#00875a','FALSE':'#de350b'}
            td['Shipment %'] = [pb(ts.iloc[i]['Shipments']/tot*100, cm.get(ss(ts.iloc[i]['Tracked']),'#666')) if ss(ts.iloc[i]['Tracked'])!='Grand Total' else '<b>100%</b>' for i in range(len(ts))]
            st.markdown(ht(td,250), unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)
        with c2:
            st.markdown('<div class="sc"><h3>Tracking by Carrier</h3>', unsafe_allow_html=True)
            tc = pv_track_carr(a); tcd = tc.copy()
            tcd['Tracked'] = [b_tr(v) for v in tc['Tracked'].tolist()]
            st.markdown(ht(tcd,300), unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

        st.markdown("#### Tracking Detail")
        f1,f2,_ = st.columns([2,3,5])
        with f1: tf = st.selectbox("Tracked",["All","TRUE","FALSE"],key="tf")
        with f2: cf = st.selectbox("Carrier",carrs,key="tc")
        d = a.copy()
        if tf!="All": d = flt(d,'Tracked',tf)
        if cf!="All": d = flt(d,'Carrier Name',cf,True)
        st.caption(f"{len(d)} shipments")
        cols = ['Carrier Name','Bill of Lading','Tracked','Connection Type','Tracking Method','Pickup Location','Destination Location','Lanes','Final Status Reason','Tracked Status','Tracking Error']
        dd = d[cols].copy()
        dd['Tracked'] = [b_tr(v) for v in dd['Tracked'].tolist()]
        dd['Tracked Status'] = [b_st(v) for v in dd['Tracked Status'].tolist()]
        st.markdown(ht(dd,500), unsafe_allow_html=True)

    # ── TAB 2 ───────────────────────────────────────────────
    with t2:
        c1,c2 = st.columns(2)
        with c1:
            st.markdown('<div class="sc"><h3>DQ-Milestone Summary</h3>', unsafe_allow_html=True)
            ms = pv_ms(a); md = ms.copy()
            cm2 = {'Full Tracked':'#00875a','Partial Tracked':'#ff991f','Tracked with 0 milestones':'#de350b'}
            md['Tracked Status'] = [b_st(v) if v!='Grand Total' else f'<b>{v}</b>' for v in ms['Tracked Status'].tolist()]
            md['Shipment %'] = [pb(ms.iloc[i]['Shipments']/tot*100, cm2.get(ss(ms.iloc[i]['Tracked Status']),'#666')) if ss(ms.iloc[i]['Tracked Status'])!='Grand Total' else '<b>100%</b>' for i in range(len(ms))]
            st.markdown(ht(md,250), unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)
        with c2:
            st.markdown('<div class="sc"><h3>P44 Root Cause Analysis</h3>', unsafe_allow_html=True)
            rc = pv_rca(a); rd = rc.copy()
            rd['P44 Analysis'] = [b_p4(v) if ss(v)!='Grand Total' else f'<b>{v}</b>' for v in rc['P44 Analysis'].tolist()]
            st.markdown(ht(rd,300), unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

        c3,c4 = st.columns(2)
        with c3:
            st.markdown('<div class="sc"><h3>Milestone by Carrier</h3>', unsafe_allow_html=True)
            mc = pv_mc(a); mcd = mc.copy()
            mcd['Milestone Missed'] = [b_mm(v) for v in mc['Milestone Missed'].tolist()]
            st.markdown(ht(mcd,300), unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)
        with c4:
            st.markdown('<div class="sc"><h3>Lane Analysis</h3>', unsafe_allow_html=True)
            la = pv_lane(a); lad = la.copy()
            lad['Milestone Missed'] = [b_mm(v) for v in la['Milestone Missed'].tolist()]
            st.markdown(ht(lad,300), unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)

        st.markdown("#### Milestone Detail")
        m1,m2,_ = st.columns([2,3,5])
        with m1: sf = st.selectbox("Tracked Status",['All','Full Tracked','Partial Tracked','Tracked with 0 milestones'],key="ms")
        with m2: mcf = st.selectbox("Carrier",carrs,key="mc")
        md2 = a.copy()
        if sf!="All": md2 = flt(md2,'Tracked Status',sf,True)
        if mcf!="All": md2 = flt(md2,'Carrier Name',mcf,True)
        st.caption(f"{len(md2)} shipments")
        mcols = ['Carrier Name','Bill of Lading','Tracked','Lanes','Pickup Location','Destination Location','Pickup Arrival Milestone (UTC)','Pickup Departure Milestone (UTC)','Final Destination Arrival Milestone (UTC)','Final Destination Departure Milestone (UTC)','Final Status Reason','Tracked Status','Milestone Achieved','Milestone Missed','p44 Analysis']
        mdd = md2[mcols].copy()
        mdd['Tracked'] = [b_tr(v) for v in mdd['Tracked'].tolist()]
        mdd['Tracked Status'] = [b_st(v) for v in mdd['Tracked Status'].tolist()]
        mdd['Milestone Missed'] = [b_mm(v) for v in mdd['Milestone Missed'].tolist()]
        mdd['p44 Analysis'] = [b_p4(v) for v in mdd['p44 Analysis'].tolist()]
        st.markdown(ht(mdd,500), unsafe_allow_html=True)

    # ── TAB 3 ───────────────────────────────────────────────
    with t3:
        st.markdown("#### Full Processed Data (Query Sheet)")
        r1,r2,_ = st.columns([2,3,5])
        with r1: rtf = st.selectbox("Tracked",["All","TRUE","FALSE"],key="rt")
        with r2: rcf = st.selectbox("Carrier",carrs,key="rcc")
        rd2 = q.copy()
        if rtf!="All": rd2 = flt(rd2,'Tracked',rtf)
        if rcf!="All": rd2 = flt(rd2,'Carrier Name',rcf,True)
        st.caption(f"{len(rd2)} shipments")
        rdd = rd2.copy()
        rdd['Tracked'] = [b_tr(v) for v in rdd['Tracked'].tolist()]
        st.markdown(ht(rdd,600), unsafe_allow_html=True)


if __name__ == '__main__':
    main()

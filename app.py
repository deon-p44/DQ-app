# ──────────────────────────────────────────────────────────────
# AUTO-INSTALL
# ──────────────────────────────────────────────────────────────
import subprocess, sys
for _p in ["openpyxl","xlsxwriter","plotly"]:
    try: __import__(_p)
    except ImportError: subprocess.check_call([sys.executable,"-m","pip","install",_p,"-q"])

import pandas as pd
try: pd.options.future.infer_string = False
except: pass

import streamlit as st
import plotly.graph_objects as go
import io, traceback as tb
from datetime import date

st.set_page_config(page_title="Data Quality Report", page_icon="📊", layout="wide", initial_sidebar_state="expanded")

# ──────────────────────────────────────────────────────────────
# CSS
# ──────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
*{font-family:'Inter',-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif}

.header{background:linear-gradient(135deg,#1a1a2e 0%,#16213e 100%);color:#fff;padding:14px 28px;display:flex;align-items:center;justify-content:space-between;box-shadow:0 2px 8px rgba(0,0,0,.2);border-radius:0;margin:-1rem -1rem 16px -1rem}
.header h1{font-size:18px;font-weight:600;letter-spacing:.5px;color:#fff;margin:0}
.header .sub{font-size:11px;opacity:.7;margin-top:1px}
.header-right{display:flex;align-items:center;gap:10px}
.header-right .date-badge{background:rgba(255,255,255,.12);padding:5px 12px;border-radius:20px;font-size:12px}
.customer-label{font-size:12px;font-weight:500;padding:3px 10px;background:rgba(255,255,255,.1);border-radius:6px}

.kpi-strip{display:flex;gap:12px;margin-bottom:16px;overflow-x:auto}
.kpi-card{flex:1;min-width:120px;padding:12px 14px;background:white;border-radius:8px;text-align:center;box-shadow:0 1px 3px rgba(0,0,0,.06);border:1px solid #edf0f3}
.kpi-value{font-size:24px;font-weight:700;color:#1a1a2e}
.kpi-label{font-size:10px;color:#6b778c;text-transform:uppercase;letter-spacing:.7px;margin-top:3px}
.kpi-card.green .kpi-value{color:#00875a}
.kpi-card.red .kpi-value{color:#de350b}
.kpi-card.orange .kpi-value{color:#ff991f}
.kpi-card.blue .kpi-value{color:#0046FF}

.section-card{background:white;border-radius:8px;box-shadow:0 1px 3px rgba(0,0,0,.06);margin-bottom:14px;overflow:hidden;border:1px solid #edf0f3}
.section-header{padding:12px 16px;border-bottom:1px solid #edf0f3;display:flex;align-items:center;justify-content:space-between}
.section-header h3{font-size:13px;font-weight:600;margin:0;color:#172b4d}
.section-count{font-size:11px;color:#6b778c}

.badge{display:inline-block;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:600;white-space:nowrap}
.badge-green{background:#e3fcef;color:#006644}
.badge-red{background:#ffebe6;color:#bf2600}
.badge-orange{background:#fff7e6;color:#974f0c}
.badge-blue{background:#deebff;color:#0747a6}

.pct-cell{display:flex;align-items:center;gap:6px}
.pct-bar{display:inline-block;height:6px;border-radius:3px;min-width:3px;vertical-align:middle;margin-right:5px}

.block-container{padding-top:0!important;max-width:1400px}
div[data-testid="stTabs"] button[role="tab"]{font-size:13px;font-weight:500;padding:10px 20px}
section[data-testid="stSidebar"]{background:#f8f9fb;border-right:1px solid #edf0f3}
section[data-testid="stSidebar"] .stSelectbox label{font-size:12px;font-weight:600;color:#6b778c;text-transform:uppercase;letter-spacing:.5px}
#MainMenu{visibility:hidden}footer{visibility:hidden}header{visibility:hidden}
</style>
""", unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────
# HELPERS
# ──────────────────────────────────────────────────────────────
def ss(v):
    if v is None: return ''
    if isinstance(v, float) and pd.isna(v): return ''
    s = str(v).strip()
    return '' if s in ('nan','None','NaT','NaN') else s

def su(v): return ss(v).upper()

def find_col(df,cands):
    lo={c.lower().strip():c for c in df.columns}
    for c in cands:
        if c.lower().strip() in lo: return lo[c.lower().strip()]
    for c in cands:
        cl=c.lower().strip()
        for k,v in lo.items():
            if cl in k: return v
    return None


# ──────────────────────────────────────────────────────────────
# PROCESS (exact Power Query logic)
# ──────────────────────────────────────────────────────────────
def process(df):
    cc=find_col(df,['Carrier Name']);bc=find_col(df,['Bill of Lading']);tc=find_col(df,['Tracked'])
    if cc is None and bc is None: st.error("Cannot find required columns.");return None,None
    keep=[]
    for i in range(len(df)):
        r=df.iloc[i];ok=False
        if cc and ss(r.get(cc,'')): ok=True
        if bc and ss(r.get(bc,'')): ok=True
        if tc and ss(r.get(tc,'')): ok=True
        keep.append(ok)
    df=df[keep].reset_index(drop=True);n=len(df)
    def gc(cands):
        c=find_col(df,cands)
        return [ss(df.iloc[i][c]) for i in range(n)] if c else ['']*n

    cn=gc(['Carrier Name']);bl=gc(['Bill of Lading']);on=gc(['Order Number']);tr=gc(['Tracked'])
    ct=gc(['Connection Type']);tm=gc(['Tracking Method']);ae=gc(['Active Equipment ID']);he=gc(['Historical Equipment ID'])
    pn=gc(['Pickup Name']);pcs=gc(['Pickup City State']);pc=gc(['Pickup Country'])
    dn=gc(['Final Destination Name']);dcs=gc(['Final Destination City State']);dc=gc(['Final Destination Country'])
    paw=gc(['Pickup Appointement Window (UTC)','Pickup Appointment Window'])
    daw=gc(['Delivery Appointement Window (UTC)','Delivery Appointment Window'])
    scr=gc(['Shipment Created (UTC)','Shipment Created'])
    tws=gc(['Tracking Window Start (UTC)','Tracking Window Start']);twe=gc(['Tracking Window End (UTC)','Tracking Window End'])
    m1r=gc(['Pickup Arrival Milestone (UTC)','Pickup Arrival Milestone'])
    m2r=gc(['Pickup Departure Milestone (UTC)','Pickup Departure Milestone'])
    m3r=gc(['Final Destination Arrival Milestone (UTC)','Final Destination Arrival'])
    m4r=gc(['Final Destination Departure Milestone (UTC)','Final Destination Departure'])
    mre=gc(['# Of Milestones received / # Of Milestones expected'])
    ur=gc(['# Updates Received']);u10=gc(['# Updates Received < 10 mins'])
    nie=gc(['Nb Intervals Expected']);nio=gc(['Nb Intervals Observed'])
    fsr=gc(['Final Status Reason']);ter=gc(['Tracking Error'])
    me1=gc(['Milestone Error 1']);me2=gc(['Milestone Error 2']);me3=gc(['Milestone Error 3'])
    def hm(v): return v not in ('','0','UNKNOWN')
    qrows,arows=[],[]
    for i in range(n):
        pl=','.join(filter(None,[pn[i],pcs[i],pc[i]]));dl=','.join(filter(None,[dn[i],dcs[i],dc[i]]))
        qrows.append({'Carrier Name':cn[i],'Bill of Lading':bl[i],'Order Number':on[i],'Tracked':tr[i],'Connection Type':ct[i],'Tracking Method':tm[i],'Active Equipment ID':ae[i],'Historical Equipment ID':he[i],'Pickup Name':pn[i],'Pickup Location':pl,'Pickup City State':pcs[i],'Pickup Country':pc[i],'Pickup Appointement Window (UTC)':paw[i],'Final Destination Name':dn[i],'Final Destination City State':dcs[i],'Final Destination Country':dc[i],'Delivery Appointement Window (UTC)':daw[i],'Shipment Created (UTC)':scr[i],'Tracking Window Start (UTC)':tws[i],'Tracking Window End (UTC)':twe[i],'Pickup Arrival Milestone (UTC)':m1r[i],'Pickup Departure Milestone (UTC)':m2r[i],'Final Destination Arrival Milestone (UTC)':m3r[i],'Final Destination Departure Milestone (UTC)':m4r[i],'# Of Milestones received / # Of Milestones expected':mre[i],'# Updates Received':ur[i],'# Updates Received < 10 mins':u10[i],'Nb Intervals Expected':nie[i],'Nb Intervals Observed':nio[i],'Final Status Reason':fsr[i],'Tracking Error':ter[i],'Milestone Error 1':me1[i],'Milestone Error 2':me2[i],'Milestone Error 3':me3[i]})
        trkd=tr[i].upper()=='TRUE';a1,a2,a3,a4=hm(m1r[i]),hm(m2r[i]),hm(m3r[i]),hm(m4r[i])
        ach=[l for f,l in [(a1,'m1'),(a2,'m2'),(a3,'m3'),(a4,'m4')] if f]
        mis=[l for f,l in [(a1,'m1'),(a2,'m2'),(a3,'m3'),(a4,'m4')] if not f]
        nc=len(ach)
        if trkd and nc==4: ts,mm,ana,p44='Full Tracked','Fully Tracked','','Full Tracked'
        elif trkd and nc>0: ts,mm,ana,p44='Partial Tracked',', '.join(mis),'','Partial Tracked'
        elif trkd: ts,mm,ana,p44='Tracked with 0 milestones','m1, m2, m3, m4','','Tracked with 0 milestones'
        else: e=ter[i];ts='Tracked with 0 milestones';mm='m1, m2, m3, m4';ana=e;p44=e if e else 'Tracked with 0 milestones'
        lane=f"{pcs[i]} -> {dcs[i]}" if pcs[i] and dcs[i] else ''
        arows.append({'Carrier Name':cn[i],'Bill of Lading':bl[i],'Order Number':on[i],'Tracked':tr[i],'Connection Type':ct[i],'Tracking Method':tm[i],'Active Equipment ID':ae[i],'Historical Equipment ID':he[i],'Lanes':lane,'Pickup Location':pl,'Destination Location':dl,'Pickup Appointement Window (UTC)':paw[i],'Delivery Appointement Window (UTC)':daw[i],'Shipment Created (UTC)':scr[i],'Tracking Window Start (UTC)':tws[i],'Tracking Window End (UTC)':twe[i],'Pickup Arrival Milestone (UTC)':m1r[i],'Pickup Departure Milestone (UTC)':m2r[i],'Final Destination Arrival Milestone (UTC)':m3r[i],'Final Destination Departure Milestone (UTC)':m4r[i],'Final Status Reason':fsr[i],'Tracked Status':ts,'Milstone Completeness':'4/4','Milestone Achieved':', '.join(ach) if ach else '','Milestone Missed':mm,'Analysis':ana,'p44 Analysis':p44,'Tracking Error':ter[i]})
    return pd.DataFrame(qrows),pd.DataFrame(arows)


# ──────────────────────────────────────────────────────────────
# EXCEL
# ──────────────────────────────────────────────────────────────
def to_xl(sheets):
    buf=io.BytesIO()
    with pd.ExcelWriter(buf,engine='xlsxwriter') as w:
        for nm,df in sheets.items():
            sn=nm[:31];df.to_excel(w,sheet_name=sn,index=False)
            ws=w.sheets[sn]
            for i,c in enumerate(df.columns):
                try: ml=max(max(len(ss(v)) for v in df.iloc[:,i].tolist()),0) if len(df)>0 else 0
                except: ml=0
                ws.set_column(i,i,min(max(ml,len(str(c)))+2,45))
    return buf.getvalue()


# ──────────────────────────────────────────────────────────────
# PIVOTS
# ──────────────────────────────────────────────────────────────
def _cnt(lst):
    d={}
    for v in lst: d[v]=d.get(v,0)+1
    return d

def pv_track(a):
    t=len(a);c=_cnt([su(v) for v in a['Tracked'].tolist()])
    r=[{'Tracked':k,'Shipments':v,'Shipment %':f"{v/t*100:.1f}%"} for k,v in sorted(c.items(),key=lambda x:-x[1])]
    r.append({'Tracked':'Grand Total','Shipments':t,'Shipment %':'100%'});return pd.DataFrame(r)

def pv_track_carr(a):
    t=len(a);c={}
    for i in range(len(a)): k=(su(a.iloc[i]['Tracked']),ss(a.iloc[i]['Carrier Name']));c[k]=c.get(k,0)+1
    return pd.DataFrame([{'Tracked':k[0],'Carrier Name':k[1],'Shipments':v,'Shipment %':f"{v/t*100:.1f}%"} for k,v in sorted(c.items(),key=lambda x:-x[1])])

def pv_ms(a):
    t=len(a);c=_cnt([ss(v) for v in a['Tracked Status'].tolist()])
    return pd.DataFrame([{'Tracked Status':s,'Shipments':c.get(s,0),'Shipment %':f"{c.get(s,0)/t*100:.1f}%"} for s in ['Full Tracked','Partial Tracked','Tracked with 0 milestones']]+[{'Tracked Status':'Grand Total','Shipments':t,'Shipment %':'100%'}])

def pv_rca(a):
    t=len(a);c=_cnt([ss(v) for v in a['p44 Analysis'].tolist()])
    r=[{'P44 Analysis':k,'Shipments':v,'Shipment %':f"{v/t*100:.1f}%"} for k,v in sorted(c.items(),key=lambda x:-x[1])]
    r.append({'P44 Analysis':'Grand Total','Shipments':t,'Shipment %':'100%'});return pd.DataFrame(r)

def pv_mc(a):
    t=len(a);c={}
    for i in range(len(a)):k=(ss(a.iloc[i]['Carrier Name']),ss(a.iloc[i]['Milestone Missed']));c[k]=c.get(k,0)+1
    return pd.DataFrame([{'Carrier Name':k[0],'Milestone Missed':k[1],'Shipments':v,'Shipment %':f"{v/t*100:.1f}%"} for k,v in sorted(c.items(),key=lambda x:-x[1])])

def pv_lane(a):
    t=len(a);c={}
    for i in range(len(a)):k=(ss(a.iloc[i]['Lanes']),ss(a.iloc[i]['Carrier Name']),ss(a.iloc[i]['Milestone Missed']));c[k]=c.get(k,0)+1
    return pd.DataFrame([{'Lanes':k[0],'Carrier Name':k[1],'Milestone Missed':k[2],'Shipments':v,'Shipment %':f"{v/t*100:.1f}%"} for k,v in sorted(c.items(),key=lambda x:-x[1])])


# ──────────────────────────────────────────────────────────────
# CHARTS (compact, clean)
# ──────────────────────────────────────────────────────────────
C={'g':'#00875a','r':'#de350b','o':'#ff991f','b':'#0046FF','d':'#1a1a2e'}
CL=dict(paper_bgcolor='rgba(0,0,0,0)',plot_bgcolor='rgba(0,0,0,0)',margin=dict(l=10,r=10,t=30,b=10),font=dict(family='Inter',size=11,color='#172b4d'))

def ch_track_donut(a):
    tl=[su(v) for v in a['Tracked'].tolist()]
    v=[sum(1 for x in tl if x=='TRUE'),sum(1 for x in tl if x!='TRUE')]
    fig=go.Figure(go.Pie(labels=['TRUE','FALSE'],values=v,hole=.6,marker=dict(colors=[C['g'],C['r']]),textinfo='label+percent',textfont=dict(size=11),pull=[0,.03]))
    fig.update_layout(**CL,height=220,showlegend=False);return fig

def ch_ms_donut(a):
    sl=[ss(v) for v in a['Tracked Status'].tolist()]
    v=[sum(1 for x in sl if x=='Full Tracked'),sum(1 for x in sl if x=='Partial Tracked'),sum(1 for x in sl if x=='Tracked with 0 milestones')]
    fig=go.Figure(go.Pie(labels=['Full','Partial','0 MS'],values=v,hole=.6,marker=dict(colors=[C['g'],C['o'],C['r']]),textinfo='label+percent',textfont=dict(size=11)))
    fig.update_layout(**CL,height=220,showlegend=False);return fig

def ch_rca(a):
    c=_cnt([ss(v) for v in a['p44 Analysis'].tolist()])
    it=sorted(c.items(),key=lambda x:-x[1])
    cols=[C['g'] if k=='Full Tracked' else C['o'] if k=='Partial Tracked' else C['b'] if 'milestones' in k.lower() else C['r'] for k,_ in it]
    fig=go.Figure(go.Bar(x=[v for _,v in it],y=[k[:40] for k,_ in it],orientation='h',marker_color=cols,text=[v for _,v in it],textposition='outside',textfont=dict(size=10)))
    h=max(200,len(it)*28+60)
    fig.update_layout(**CL,height=h,yaxis=dict(autorange='reversed',tickfont=dict(size=10)),xaxis=dict(showgrid=True,gridcolor='#f0f2f5',zeroline=False),bargap=.3);return fig

def ch_carrier(a):
    c={}
    for i in range(len(a)):
        cn=ss(a.iloc[i]['Carrier Name']);ts=ss(a.iloc[i]['Tracked Status'])
        if cn not in c:c[cn]={'Full Tracked':0,'Partial Tracked':0,'Tracked with 0 milestones':0}
        if ts in c[cn]:c[cn][ts]+=1
    top=sorted(c.keys(),key=lambda x:-sum(c[x].values()))[:10]
    fig=go.Figure()
    for s,cl in [('Full Tracked',C['g']),('Partial Tracked',C['o']),('Tracked with 0 milestones',C['r'])]:
        fig.add_trace(go.Bar(name=s,y=[cn[:25] for cn in top],x=[c[cn].get(s,0) for cn in top],orientation='h',marker_color=cl,textfont=dict(size=10)))
    h=max(200,len(top)*28+80)
    fig.update_layout(**CL,barmode='stack',height=h,yaxis=dict(autorange='reversed',tickfont=dict(size=10)),xaxis=dict(showgrid=True,gridcolor='#f0f2f5',zeroline=False),legend=dict(orientation='h',y=-0.2,font=dict(size=10)));return fig


# ──────────────────────────────────────────────────────────────
# BADGES & HTML TABLE
# ──────────────────────────────────────────────────────────────
def bg(t,c): return f'<span class="badge badge-{c}">{t}</span>'
def b_tr(v): return bg(v,'green') if su(v)=='TRUE' else bg(v,'red')
def b_st(v): return bg(v,'green') if v=='Full Tracked' else bg(v,'orange') if v=='Partial Tracked' else bg(v,'red')
def b_mm(v): return bg(v,'green') if v=='Fully Tracked' else bg(v,'orange')
def b_p4(v):
    if v=='Full Tracked':return bg(v,'green')
    if v=='Partial Tracked':return bg(v,'orange')
    if 'milestones' in ss(v).lower():return bg(v,'blue')
    return bg(v,'red')
def pb(p,c='#00875a'): return f'<div class="pct-cell"><span class="pct-bar" style="width:{max(p,2):.0f}px;background:{c}"></span>{p:.1f}%</div>'

def ht(df,mh=350):
    h=f'<div style="max-height:{mh}px;overflow:auto"><table style="width:100%;border-collapse:collapse;font-size:12px">'
    h+='<thead><tr>'
    for c in df.columns:
        h+=f'<th style="background:#f8f9fb;padding:8px 12px;text-align:left;font-weight:600;font-size:11px;text-transform:uppercase;letter-spacing:.4px;color:#6b778c;border-bottom:2px solid #dfe1e6;white-space:nowrap;position:sticky;top:0;z-index:1">{c}</th>'
    h+='</tr></thead><tbody>'
    for i in range(len(df)):
        r=df.iloc[i];tot='Total' in ss(r.iloc[0])
        s='background:#f0f2f5;font-weight:700;border-top:2px solid #dfe1e6' if tot else ''
        h+=f'<tr style="{s}">'
        for v in r: h+=f'<td style="padding:7px 12px;border-bottom:1px solid #f0f0f0;white-space:nowrap">{ss(v)}</td>'
        h+='</tr>'
    h+='</tbody></table></div>';return h

def flt(df,col,val,exact=False):
    m=[ss(df.iloc[i][col])==val if exact else su(df.iloc[i][col])==val.upper() for i in range(len(df))]
    return df[m].reset_index(drop=True)


# ──────────────────────────────────────────────────────────────
# MAIN
# ──────────────────────────────────────────────────────────────
def main():
    today=date.today().strftime('%b %d, %Y')
    cust=st.session_state.get('cust','')
    st.markdown(f'<div class="header"><div><h1>Data Quality Report</h1><div class="sub">project44 Visibility Platform</div></div><div class="header-right"><span class="customer-label">{cust}</span><span class="date-badge">{today}</span></div></div>',unsafe_allow_html=True)

    if 'ok' not in st.session_state: st.session_state.ok=False
    up=st.file_uploader("Upload your weekly data export (.xlsx)",type=['xlsx','xls','csv'],label_visibility="collapsed")

    if up and not st.session_state.ok:
        with st.spinner("Processing data…"):
            try:
                if up.name.endswith('.csv'):raw=pd.read_csv(up)
                else:
                    xls=pd.ExcelFile(up,engine='openpyxl')
                    sh='Data' if 'Data' in xls.sheet_names else xls.sheet_names[0]
                    raw=pd.read_excel(xls,sheet_name=sh,engine='openpyxl')
                q,a=process(raw)
                if q is not None:
                    st.session_state.q=q;st.session_state.a=a;st.session_state.ok=True
                    cc=find_col(raw,['Customer Tenant Name'])
                    st.session_state.cust=ss(raw[cc].dropna().iloc[0]) if cc and len(raw[cc].dropna())>0 else ''
                    st.rerun()
            except Exception as e:st.error(f"Error: {e}");st.code(tb.format_exc());return

    if not st.session_state.ok:
        st.markdown("""
        <div style="text-align:center;padding:80px 20px">
            <div style="font-size:48px;margin-bottom:16px">📊</div>
            <h2 style="font-size:18px;margin-bottom:8px">Upload Data Quality Report</h2>
            <p style="font-size:13px;color:#6b778c">Drag & drop your .xlsx file above or click to browse.<br>The system will automatically process the data and generate Tracking and Milestone reports.</p>
        </div>""",unsafe_allow_html=True)
        return

    q=st.session_state.q;a_full=st.session_state.a

    # ── SIDEBAR: GLOBAL FILTERS ─────────────────────────────
    with st.sidebar:
        st.markdown("### 🔍 Filters")
        st.caption("Applied across all tabs")
        carrs_all=sorted(set(ss(v) for v in a_full['Carrier Name'].tolist() if ss(v)))
        f_tracked=st.selectbox("Tracked",["All","TRUE","FALSE"],key="gf_t")
        f_status=st.selectbox("Tracked Status",["All","Full Tracked","Partial Tracked","Tracked with 0 milestones"],key="gf_s")
        f_carrier=st.selectbox("Carrier",["All"]+carrs_all,key="gf_c")
        lanes_all=sorted(set(ss(v) for v in a_full['Lanes'].tolist() if ss(v)))
        f_lane=st.selectbox("Lane",["All"]+lanes_all,key="gf_l")
        st.divider()
        if st.button("↺ Upload New File",use_container_width=True):
            for k in ['ok','q','a','cust']:st.session_state.pop(k,None)
            st.rerun()
        st.divider()
        st.markdown("##### 📥 Export")
        st.download_button("⬇ Export All to Excel",to_xl({'Query':q,'Data Analysis':a_full,'DQ-Tracking':pv_track(a_full),'Tracking Carrier':pv_track_carr(a_full),'DQ-Milestone':pv_ms(a_full),'P44 RCA':pv_rca(a_full),'Lane Analysis':pv_lane(a_full),'MS Carrier':pv_mc(a_full)}),"DQ-Report-Full.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True,type="primary")
        st.download_button("⬇ Tracking Report",to_xl({'DQ-Tracking':pv_track(a_full),'Carrier':pv_track_carr(a_full),'Detail':a_full[['Carrier Name','Bill of Lading','Tracked','Connection Type','Tracking Method','Lanes','Pickup Location','Destination Location','Final Status Reason','Tracked Status','Tracking Error']]}),"DQ-Tracking.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
        st.download_button("⬇ Milestone Report",to_xl({'DQ-Milestone':pv_ms(a_full),'P44 RCA':pv_rca(a_full),'MS Carrier':pv_mc(a_full),'Lane':pv_lane(a_full),'Detail':a_full}),"DQ-Milestone.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)

    # Apply global filters
    a=a_full.copy()
    if f_tracked!="All": a=flt(a,'Tracked',f_tracked)
    if f_status!="All": a=flt(a,'Tracked Status',f_status,True)
    if f_carrier!="All": a=flt(a,'Carrier Name',f_carrier,True)
    if f_lane!="All": a=flt(a,'Lanes',f_lane,True)
    # Also filter query_df
    qf=q.copy()
    if f_tracked!="All": qf=flt(qf,'Tracked',f_tracked)
    if f_carrier!="All": qf=flt(qf,'Carrier Name',f_carrier,True)

    # ── KPIs (from filtered data) ───────────────────────────
    tot=len(a)
    if tot==0:
        st.warning("No data matches the current filters. Adjust filters in the sidebar.");return
    tl=[su(v) for v in a['Tracked'].tolist()];sl=[ss(v) for v in a['Tracked Status'].tolist()]
    tt=sum(1 for v in tl if v=='TRUE');ft=sum(1 for v in sl if v=='Full Tracked')
    pt=sum(1 for v in sl if v=='Partial Tracked');zm=sum(1 for v in sl if v=='Tracked with 0 milestones')
    tp=tt/tot*100 if tot else 0;fp=ft/tot*100 if tot else 0

    st.markdown(f"""<div class="kpi-strip">
    <div class="kpi-card blue"><div class="kpi-value">{tot:,}</div><div class="kpi-label">Total Shipments</div></div>
    <div class="kpi-card green"><div class="kpi-value">{tp:.1f}%</div><div class="kpi-label">Tracked (TRUE)</div></div>
    <div class="kpi-card green"><div class="kpi-value">{ft:,}</div><div class="kpi-label">Full Tracked</div></div>
    <div class="kpi-card orange"><div class="kpi-value">{pt:,}</div><div class="kpi-label">Partial Tracked</div></div>
    <div class="kpi-card red"><div class="kpi-value">{zm:,}</div><div class="kpi-label">0 Milestones</div></div>
    <div class="kpi-card"><div class="kpi-value">{fp:.1f}%</div><div class="kpi-label">Full Track Rate</div></div>
    </div>""",unsafe_allow_html=True)

    # ── TABS ────────────────────────────────────────────────
    tab1,tab2,tab3=st.tabs(["📡 Tracking Data","🎯 Milestone Data","📋 Processed Data"])

    # ═══════════════════ TAB 1 ═════════════════════════════
    with tab1:
        c1,c2=st.columns(2)
        with c1:
            st.markdown('<div class="section-card"><div class="section-header"><h3>Tracking Rate</h3></div><div class="section-body" style="padding:8px">',unsafe_allow_html=True)
            st.plotly_chart(ch_track_donut(a),use_container_width=True,config={'displayModeBar':False})
            st.markdown('</div></div>',unsafe_allow_html=True)
        with c2:
            st.markdown('<div class="section-card"><div class="section-header"><h3>Top Carriers by Status</h3></div><div class="section-body" style="padding:8px">',unsafe_allow_html=True)
            st.plotly_chart(ch_carrier(a),use_container_width=True,config={'displayModeBar':False})
            st.markdown('</div></div>',unsafe_allow_html=True)

        c3,c4=st.columns(2)
        with c3:
            st.markdown('<div class="section-card"><div class="section-header"><h3>DQ-Tracking Summary</h3></div><div class="section-body">',unsafe_allow_html=True)
            ts=pv_track(a);td=ts.copy()
            cm={'TRUE':'#00875a','FALSE':'#de350b'}
            td['Tracked']=[b_tr(v) if v!='Grand Total' else f'<b>{v}</b>' for v in ts['Tracked'].tolist()]
            td['Shipment %']=[pb(ts.iloc[i]['Shipments']/tot*100,cm.get(ss(ts.iloc[i]['Tracked']),'#666')) if ss(ts.iloc[i]['Tracked'])!='Grand Total' else '<b>100%</b>' for i in range(len(ts))]
            st.markdown(ht(td,200),unsafe_allow_html=True)
            st.markdown('</div></div>',unsafe_allow_html=True)
        with c4:
            st.markdown('<div class="section-card"><div class="section-header"><h3>Tracking by Carrier</h3></div><div class="section-body">',unsafe_allow_html=True)
            tc=pv_track_carr(a);tcd=tc.copy()
            tcd['Tracked']=[b_tr(v) for v in tc['Tracked'].tolist()]
            st.markdown(ht(tcd,250),unsafe_allow_html=True)
            st.markdown('</div></div>',unsafe_allow_html=True)

        st.markdown(f'<div class="section-card"><div class="section-header"><h3>Tracking Detail</h3><span class="section-count">{len(a)} shipments</span></div><div class="section-body">',unsafe_allow_html=True)
        cols=['Carrier Name','Bill of Lading','Tracked','Connection Type','Tracking Method','Pickup Location','Destination Location','Lanes','Final Status Reason','Tracked Status','Tracking Error']
        dd=a[cols].copy()
        dd['Tracked']=[b_tr(v) for v in dd['Tracked'].tolist()]
        dd['Tracked Status']=[b_st(v) for v in dd['Tracked Status'].tolist()]
        st.markdown(ht(dd,450),unsafe_allow_html=True)
        st.markdown('</div></div>',unsafe_allow_html=True)
        st.download_button("⬇ Download Filtered Tracking Data",to_xl({'Tracking':a[cols]}),"Tracking-Filtered.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",key="dl_t")

    # ═══════════════════ TAB 2 ═════════════════════════════
    with tab2:
        c1,c2=st.columns(2)
        with c1:
            st.markdown('<div class="section-card"><div class="section-header"><h3>Milestone Completeness</h3></div><div class="section-body" style="padding:8px">',unsafe_allow_html=True)
            st.plotly_chart(ch_ms_donut(a),use_container_width=True,config={'displayModeBar':False})
            st.markdown('</div></div>',unsafe_allow_html=True)
        with c2:
            st.markdown('<div class="section-card"><div class="section-header"><h3>P44 Root Cause Analysis</h3></div><div class="section-body" style="padding:8px">',unsafe_allow_html=True)
            st.plotly_chart(ch_rca(a),use_container_width=True,config={'displayModeBar':False})
            st.markdown('</div></div>',unsafe_allow_html=True)

        c3,c4=st.columns(2)
        with c3:
            st.markdown('<div class="section-card"><div class="section-header"><h3>DQ-Milestone Summary</h3></div><div class="section-body">',unsafe_allow_html=True)
            ms=pv_ms(a);md=ms.copy()
            cm2={'Full Tracked':'#00875a','Partial Tracked':'#ff991f','Tracked with 0 milestones':'#de350b'}
            md['Tracked Status']=[b_st(v) if v!='Grand Total' else f'<b>{v}</b>' for v in ms['Tracked Status'].tolist()]
            md['Shipment %']=[pb(ms.iloc[i]['Shipments']/tot*100,cm2.get(ss(ms.iloc[i]['Tracked Status']),'#666')) if ss(ms.iloc[i]['Tracked Status'])!='Grand Total' else '<b>100%</b>' for i in range(len(ms))]
            st.markdown(ht(md,200),unsafe_allow_html=True)
            st.markdown('</div></div>',unsafe_allow_html=True)
        with c4:
            st.markdown('<div class="section-card"><div class="section-header"><h3>P44 Root Cause Analysis</h3></div><div class="section-body">',unsafe_allow_html=True)
            rc=pv_rca(a);rd=rc.copy()
            rd['P44 Analysis']=[b_p4(v) if ss(v)!='Grand Total' else f'<b>{v}</b>' for v in rc['P44 Analysis'].tolist()]
            st.markdown(ht(rd,250),unsafe_allow_html=True)
            st.markdown('</div></div>',unsafe_allow_html=True)

        c5,c6=st.columns(2)
        with c5:
            st.markdown('<div class="section-card"><div class="section-header"><h3>Milestone by Carrier</h3></div><div class="section-body">',unsafe_allow_html=True)
            mc=pv_mc(a);mcd=mc.copy()
            mcd['Milestone Missed']=[b_mm(v) for v in mc['Milestone Missed'].tolist()]
            st.markdown(ht(mcd,250),unsafe_allow_html=True)
            st.markdown('</div></div>',unsafe_allow_html=True)
        with c6:
            st.markdown('<div class="section-card"><div class="section-header"><h3>Lane Analysis</h3></div><div class="section-body">',unsafe_allow_html=True)
            la=pv_lane(a);lad=la.copy()
            lad['Milestone Missed']=[b_mm(v) for v in la['Milestone Missed'].tolist()]
            st.markdown(ht(lad,250),unsafe_allow_html=True)
            st.markdown('</div></div>',unsafe_allow_html=True)

        mcols=['Carrier Name','Bill of Lading','Tracked','Lanes','Pickup Location','Destination Location','Pickup Arrival Milestone (UTC)','Pickup Departure Milestone (UTC)','Final Destination Arrival Milestone (UTC)','Final Destination Departure Milestone (UTC)','Final Status Reason','Tracked Status','Milestone Achieved','Milestone Missed','p44 Analysis']
        st.markdown(f'<div class="section-card"><div class="section-header"><h3>Milestone Detail</h3><span class="section-count">{len(a)} shipments</span></div><div class="section-body">',unsafe_allow_html=True)
        mdd=a[mcols].copy()
        mdd['Tracked']=[b_tr(v) for v in mdd['Tracked'].tolist()]
        mdd['Tracked Status']=[b_st(v) for v in mdd['Tracked Status'].tolist()]
        mdd['Milestone Missed']=[b_mm(v) for v in mdd['Milestone Missed'].tolist()]
        mdd['p44 Analysis']=[b_p4(v) for v in mdd['p44 Analysis'].tolist()]
        st.markdown(ht(mdd,450),unsafe_allow_html=True)
        st.markdown('</div></div>',unsafe_allow_html=True)
        st.download_button("⬇ Download Filtered Milestone Data",to_xl({'Milestone':a[mcols]}),"Milestone-Filtered.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",key="dl_m")

    # ═══════════════════ TAB 3 ═════════════════════════════
    with tab3:
        st.markdown(f'<div class="section-card"><div class="section-header"><h3>Full Processed Data (Query Sheet)</h3><span class="section-count">{len(qf)} shipments</span></div><div class="section-body">',unsafe_allow_html=True)
        rdd=qf.copy()
        rdd['Tracked']=[b_tr(v) for v in rdd['Tracked'].tolist()]
        st.markdown(ht(rdd,550),unsafe_allow_html=True)
        st.markdown('</div></div>',unsafe_allow_html=True)
        st.download_button("⬇ Download Filtered Data",to_xl({'Processed':qf}),"Processed-Filtered.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",key="dl_r")


if __name__=='__main__':
    main()

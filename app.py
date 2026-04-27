# ──────────────────────────────────────────────────────────────
import subprocess,sys
for _p in ["openpyxl","xlsxwriter","plotly"]:
    try:__import__(_p)
    except ImportError:subprocess.check_call([sys.executable,"-m","pip","install",_p,"-q"])
import pandas as pd
try:pd.options.future.infer_string=False
except:pass
import streamlit as st
import plotly.graph_objects as go
import io,traceback as tb
from datetime import date

st.set_page_config(page_title="Data Quality Report",page_icon="📊",layout="wide",initial_sidebar_state="expanded")

# ──────────────────────────────────────────────────────────────
# CSS — modern, clean, p44 themed
# ──────────────────────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
*{font-family:'Inter',-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif}

/* Header */
.hdr{background:linear-gradient(135deg,#0f1629 0%,#1a2744 60%,#0046FF 100%);color:#fff;padding:18px 32px;display:flex;align-items:center;justify-content:space-between;border-radius:12px;margin-bottom:20px;box-shadow:0 4px 20px rgba(0,70,255,.15)}
.hdr h1{font-size:20px;font-weight:700;color:#fff;margin:0;letter-spacing:.3px}
.hdr .sub{font-size:11px;opacity:.6;margin-top:2px;font-weight:400}
.hdr-right{display:flex;align-items:center;gap:10px}
.hdr-badge{background:rgba(255,255,255,.1);padding:5px 14px;border-radius:20px;font-size:12px;backdrop-filter:blur(4px);border:1px solid rgba(255,255,255,.08)}
.hdr-cust{font-size:12px;font-weight:500;padding:4px 12px;background:rgba(255,255,255,.08);border-radius:6px;border:1px solid rgba(255,255,255,.06)}

/* KPI Strip */
.kpi-row{display:grid;grid-template-columns:repeat(6,1fr);gap:10px;margin-bottom:20px}
@media(max-width:900px){.kpi-row{grid-template-columns:repeat(3,1fr)}}
.kpi{background:#fff;border-radius:10px;padding:16px;text-align:center;box-shadow:0 1px 4px rgba(0,0,0,.04);border:1px solid #edf0f3;position:relative;overflow:hidden}
.kpi::before{content:'';position:absolute;top:0;left:0;right:0;height:3px;border-radius:10px 10px 0 0}
.kpi.k-blue::before{background:linear-gradient(90deg,#0046FF,#4d8bff)}
.kpi.k-green::before{background:linear-gradient(90deg,#00875a,#36b37e)}
.kpi.k-emerald::before{background:linear-gradient(90deg,#006644,#00875a)}
.kpi.k-orange::before{background:linear-gradient(90deg,#ff8b00,#ff991f)}
.kpi.k-red::before{background:linear-gradient(90deg,#bf2600,#de350b)}
.kpi.k-purple::before{background:linear-gradient(90deg,#5243aa,#6554c0)}
.kpi .kv{font-size:26px;font-weight:800;margin-top:4px}
.kpi .kl{font-size:9.5px;color:#7a869a;text-transform:uppercase;letter-spacing:.8px;margin-top:6px;font-weight:600}
.kpi .kd{font-size:10px;color:#97a0af;margin-top:2px;font-weight:400}
.k-blue .kv{color:#0046FF}.k-green .kv{color:#00875a}.k-emerald .kv{color:#006644}
.k-orange .kv{color:#ff8b00}.k-red .kv{color:#bf2600}.k-purple .kv{color:#5243aa}

/* Cards */
.card{background:#fff;border-radius:10px;box-shadow:0 1px 4px rgba(0,0,0,.04);margin-bottom:14px;overflow:hidden;border:1px solid #edf0f3}
.card-h{padding:14px 18px;border-bottom:1px solid #edf0f3;display:flex;align-items:center;justify-content:space-between}
.card-h h3{font-size:13px;font-weight:700;margin:0;color:#172b4d;letter-spacing:.2px}
.card-h .cnt{font-size:11px;color:#7a869a;font-weight:500}
.card-b{padding:0}

/* Badges */
.b{display:inline-block;padding:3px 9px;border-radius:10px;font-size:10.5px;font-weight:600;white-space:nowrap;letter-spacing:.2px}
.b-g{background:#e3fcef;color:#006644}.b-r{background:#ffebe6;color:#bf2600}
.b-o{background:#fff7e6;color:#974f0c}.b-bl{background:#deebff;color:#0747a6}
.b-pu{background:#eae6ff;color:#403294}

/* Pct bar */
.pc{display:flex;align-items:center;gap:6px}
.pb{display:inline-block;height:6px;border-radius:3px;min-width:3px}

/* Sidebar */
section[data-testid="stSidebar"]{background:linear-gradient(180deg,#f8f9fb,#fff);border-right:1px solid #edf0f3}
section[data-testid="stSidebar"] .stSelectbox label{font-size:11px;font-weight:700;color:#7a869a;text-transform:uppercase;letter-spacing:.6px}
section[data-testid="stSidebar"] h3{font-size:14px;color:#172b4d}

/* Streamlit */
.block-container{padding-top:0!important;max-width:1440px}
div[data-testid="stTabs"] button[role="tab"]{font-size:13px;font-weight:600;padding:12px 22px}
#MainMenu{visibility:hidden}footer{visibility:hidden}header{visibility:hidden}
</style>
""",unsafe_allow_html=True)


# ──────────────────────────────────────────────────────────────
# HELPERS
# ──────────────────────────────────────────────────────────────
def ss(v):
    if v is None:return ''
    if isinstance(v,float) and pd.isna(v):return ''
    s=str(v).strip()
    return '' if s in ('nan','None','NaT','NaN') else s
def su(v):return ss(v).upper()
def find_col(df,cands):
    lo={c.lower().strip():c for c in df.columns}
    for c in cands:
        if c.lower().strip() in lo:return lo[c.lower().strip()]
    for c in cands:
        cl=c.lower().strip()
        for k,v in lo.items():
            if cl in k:return v
    return None


# ──────────────────────────────────────────────────────────────
# PROCESS
# ──────────────────────────────────────────────────────────────
def process(df):
    cc=find_col(df,['Carrier Name']);bc=find_col(df,['Bill of Lading']);tc=find_col(df,['Tracked'])
    if cc is None and bc is None:st.error("Cannot find required columns.");return None,None
    keep=[]
    for i in range(len(df)):
        r=df.iloc[i];ok=False
        if cc and ss(r.get(cc,'')):ok=True
        if bc and ss(r.get(bc,'')):ok=True
        if tc and ss(r.get(tc,'')):ok=True
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
    def hm(v):return v not in ('','0','UNKNOWN')
    qrows,arows=[],[]
    for i in range(n):
        pl=','.join(filter(None,[pn[i],pcs[i],pc[i]]));dl=','.join(filter(None,[dn[i],dcs[i],dc[i]]))
        qrows.append({'Carrier Name':cn[i],'Bill of Lading':bl[i],'Order Number':on[i],'Tracked':tr[i],'Connection Type':ct[i],'Tracking Method':tm[i],'Active Equipment ID':ae[i],'Historical Equipment ID':he[i],'Pickup Name':pn[i],'Pickup Location':pl,'Pickup City State':pcs[i],'Pickup Country':pc[i],'Pickup Appointement Window (UTC)':paw[i],'Final Destination Name':dn[i],'Final Destination City State':dcs[i],'Final Destination Country':dc[i],'Delivery Appointement Window (UTC)':daw[i],'Shipment Created (UTC)':scr[i],'Tracking Window Start (UTC)':tws[i],'Tracking Window End (UTC)':twe[i],'Pickup Arrival Milestone (UTC)':m1r[i],'Pickup Departure Milestone (UTC)':m2r[i],'Final Destination Arrival Milestone (UTC)':m3r[i],'Final Destination Departure Milestone (UTC)':m4r[i],'# Of Milestones received / # Of Milestones expected':mre[i],'# Updates Received':ur[i],'# Updates Received < 10 mins':u10[i],'Nb Intervals Expected':nie[i],'Nb Intervals Observed':nio[i],'Final Status Reason':fsr[i],'Tracking Error':ter[i],'Milestone Error 1':me1[i],'Milestone Error 2':me2[i],'Milestone Error 3':me3[i]})
        trkd=tr[i].upper()=='TRUE';a1,a2,a3,a4=hm(m1r[i]),hm(m2r[i]),hm(m3r[i]),hm(m4r[i])
        ach=[l for f,l in [(a1,'m1'),(a2,'m2'),(a3,'m3'),(a4,'m4')] if f]
        mis=[l for f,l in [(a1,'m1'),(a2,'m2'),(a3,'m3'),(a4,'m4')] if not f]
        nc=len(ach)
        if trkd and nc==4:ts,mm,ana,p44='Full Tracked','Fully Tracked','','Full Tracked'
        elif trkd and nc>0:ts,mm,ana,p44='Partial Tracked',', '.join(mis),'','Partial Tracked'
        elif trkd:ts,mm,ana,p44='Tracked with 0 milestones','m1, m2, m3, m4','','Tracked with 0 milestones'
        else:e=ter[i];ts='Tracked with 0 milestones';mm='m1, m2, m3, m4';ana=e;p44=e if e else 'Tracked with 0 milestones'
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
            sn=nm[:31];df.to_excel(w,sheet_name=sn,index=False);ws=w.sheets[sn]
            for i,c in enumerate(df.columns):
                try:ml=max(max(len(ss(v)) for v in df.iloc[:,i].tolist()),0) if len(df)>0 else 0
                except:ml=0
                ws.set_column(i,i,min(max(ml,len(str(c)))+2,45))
    return buf.getvalue()


# ──────────────────────────────────────────────────────────────
# PIVOTS
# ──────────────────────────────────────────────────────────────
def _cnt(lst):
    d={}
    for v in lst:d[v]=d.get(v,0)+1
    return d
def pv_track(a):
    t=len(a);c=_cnt([su(v) for v in a['Tracked'].tolist()])
    r=[{'Tracked':k,'Shipments':v,'Shipment %':f"{v/t*100:.1f}%"} for k,v in sorted(c.items(),key=lambda x:-x[1])]
    r.append({'Tracked':'Grand Total','Shipments':t,'Shipment %':'100%'});return pd.DataFrame(r)
def pv_track_carr(a):
    t=len(a);c={}
    for i in range(len(a)):k=(su(a.iloc[i]['Tracked']),ss(a.iloc[i]['Carrier Name']));c[k]=c.get(k,0)+1
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
# CHARTS
# ──────────────────────────────────────────────────────────────
CC={'g':'#00875a','r':'#de350b','o':'#ff8b00','b':'#0046FF','p':'#6554c0'}
CL=dict(paper_bgcolor='rgba(0,0,0,0)',plot_bgcolor='rgba(0,0,0,0)',margin=dict(l=8,r=8,t=8,b=8),font=dict(family='Inter',size=11,color='#172b4d'))

def ch_donut(labels,values,colors):
    fig=go.Figure(go.Pie(labels=labels,values=values,hole=.62,marker=dict(colors=colors,line=dict(color='#fff',width=2)),textinfo='label+percent',textfont=dict(size=10),pull=[0]+[.02]*(len(labels)-1)))
    fig.update_layout(**CL,height=200,showlegend=False);return fig

def ch_hbar(items,colors):
    fig=go.Figure(go.Bar(x=[v for _,v in items],y=[k[:35] for k,_ in items],orientation='h',marker_color=colors,text=[f'{v}' for _,v in items],textposition='outside',textfont=dict(size=10)))
    h=max(180,len(items)*26+50)
    fig.update_layout(**CL,height=h,yaxis=dict(autorange='reversed',tickfont=dict(size=9.5)),xaxis=dict(showgrid=True,gridcolor='#f4f5f7',zeroline=False),bargap=.35);return fig

def ch_track(a):
    tl=[su(v) for v in a['Tracked'].tolist()]
    return ch_donut(['Tracked','Untracked'],[sum(1 for x in tl if x=='TRUE'),sum(1 for x in tl if x!='TRUE')],[CC['g'],CC['r']])

def ch_ms(a):
    sl=[ss(v) for v in a['Tracked Status'].tolist()]
    return ch_donut(['Full','Partial','Untracked'],[sum(1 for x in sl if x=='Full Tracked'),sum(1 for x in sl if x=='Partial Tracked'),sum(1 for x in sl if x=='Tracked with 0 milestones')],[CC['g'],CC['o'],CC['r']])

def ch_rca(a):
    c=_cnt([ss(v) for v in a['p44 Analysis'].tolist()])
    it=sorted(c.items(),key=lambda x:-x[1])
    cols=[CC['g'] if k=='Full Tracked' else CC['o'] if k=='Partial Tracked' else CC['b'] if 'milestones' in k.lower() else CC['r'] for k,_ in it]
    return ch_hbar(it,cols)

def ch_carrier(a):
    c={}
    for i in range(len(a)):
        cn=ss(a.iloc[i]['Carrier Name']);ts=ss(a.iloc[i]['Tracked Status'])
        if cn not in c:c[cn]={'Full Tracked':0,'Partial Tracked':0,'Tracked with 0 milestones':0}
        if ts in c[cn]:c[cn][ts]+=1
    top=sorted(c.keys(),key=lambda x:-sum(c[x].values()))[:10]
    fig=go.Figure()
    for s,cl in [('Full Tracked',CC['g']),('Partial Tracked',CC['o']),('Tracked with 0 milestones',CC['r'])]:
        fig.add_trace(go.Bar(name=s,y=[cn[:22] for cn in top],x=[c[cn].get(s,0) for cn in top],orientation='h',marker_color=cl))
    h=max(200,len(top)*26+60)
    fig.update_layout(**CL,barmode='stack',height=h,yaxis=dict(autorange='reversed',tickfont=dict(size=9.5)),xaxis=dict(showgrid=True,gridcolor='#f4f5f7',zeroline=False),legend=dict(orientation='h',y=-0.25,font=dict(size=9.5)));return fig


# ──────────────────────────────────────────────────────────────
# BADGES & TABLE
# ──────────────────────────────────────────────────────────────
def bg(t,c):return f'<span class="b b-{c}">{t}</span>'
def b_tr(v):return bg(v,'g') if su(v)=='TRUE' else bg(v,'r')
def b_st(v):return bg(v,'g') if v=='Full Tracked' else bg(v,'o') if v=='Partial Tracked' else bg(v,'r')
def b_mm(v):return bg(v,'g') if v=='Fully Tracked' else bg(v,'o')
def b_p4(v):
    if v=='Full Tracked':return bg(v,'g')
    if v=='Partial Tracked':return bg(v,'o')
    if 'milestones' in ss(v).lower():return bg(v,'bl')
    return bg(v,'r')
def pbar(p,c='#00875a'):return f'<div class="pc"><span class="pb" style="width:{max(p,2):.0f}px;background:{c}"></span>{p:.1f}%</div>'

def ht(df,mh=350):
    h=f'<div style="max-height:{mh}px;overflow:auto;border-radius:0 0 8px 8px"><table style="width:100%;border-collapse:collapse;font-size:12px">'
    h+='<thead><tr>'
    for c in df.columns:h+=f'<th style="background:#f8f9fb;padding:9px 12px;text-align:left;font-weight:700;font-size:10px;text-transform:uppercase;letter-spacing:.5px;color:#7a869a;border-bottom:2px solid #edf0f3;white-space:nowrap;position:sticky;top:0;z-index:1">{c}</th>'
    h+='</tr></thead><tbody>'
    for i in range(len(df)):
        r=df.iloc[i];tot='Total' in ss(r.iloc[0])
        s='background:#f4f5f7;font-weight:700;border-top:2px solid #dfe1e6' if tot else ''
        h+=f'<tr style="{s}">'
        for v in r:h+=f'<td style="padding:8px 12px;border-bottom:1px solid #f4f5f7;white-space:nowrap">{ss(v)}</td>'
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
    st.markdown(f'<div class="hdr"><div><h1>Data Quality Report</h1><div class="sub">project44 Visibility Platform</div></div><div class="hdr-right"><span class="hdr-cust">{cust}</span><span class="hdr-badge">{today}</span></div></div>',unsafe_allow_html=True)

    if 'ok' not in st.session_state:st.session_state.ok=False
    up=st.file_uploader("Upload your weekly data export (.xlsx)",type=['xlsx','xls','csv'],label_visibility="collapsed")

    if up and not st.session_state.ok:
        with st.spinner("Processing data…"):
            try:
                if up.name.endswith('.csv'):raw=pd.read_csv(up)
                else:
                    xls=pd.ExcelFile(up,engine='openpyxl');sh='Data' if 'Data' in xls.sheet_names else xls.sheet_names[0]
                    raw=pd.read_excel(xls,sheet_name=sh,engine='openpyxl')
                q,a=process(raw)
                if q is not None:
                    st.session_state.q=q;st.session_state.a=a;st.session_state.ok=True
                    cc=find_col(raw,['Customer Tenant Name'])
                    st.session_state.cust=ss(raw[cc].dropna().iloc[0]) if cc and len(raw[cc].dropna())>0 else ''
                    st.rerun()
            except Exception as e:st.error(f"Error: {e}");st.code(tb.format_exc());return

    if not st.session_state.ok:
        st.markdown('<div style="text-align:center;padding:60px 20px"><div style="font-size:48px;margin-bottom:12px">📊</div><h2 style="font-size:18px;margin-bottom:6px;color:#172b4d">Upload Data Quality Report</h2><p style="font-size:13px;color:#7a869a;max-width:420px;margin:0 auto">Drop your weekly .xlsx export above. The system will process and generate Tracking & Milestone reports automatically.</p></div>',unsafe_allow_html=True)
        return

    q=st.session_state.q;a_full=st.session_state.a

    # ── SIDEBAR ─────────────────────────────────────────────
    with st.sidebar:
        st.markdown("### 🔍 Global Filters")
        st.caption("Applied to all tabs, KPIs, and charts")
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
        st.markdown("##### 📥 Exports")
        st.download_button("⬇ Export All to Excel",to_xl({'Query':q,'Data Analysis':a_full,'DQ-Tracking':pv_track(a_full),'Tracking Carrier':pv_track_carr(a_full),'DQ-Milestone':pv_ms(a_full),'P44 RCA':pv_rca(a_full),'Lane Analysis':pv_lane(a_full),'MS Carrier':pv_mc(a_full)}),"DQ-Report-Full.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True,type="primary")
        st.download_button("⬇ Tracking Report",to_xl({'DQ-Tracking':pv_track(a_full),'Carrier':pv_track_carr(a_full),'Detail':a_full[['Carrier Name','Bill of Lading','Tracked','Connection Type','Tracking Method','Lanes','Pickup Location','Destination Location','Final Status Reason','Tracked Status','Tracking Error']]}),"DQ-Tracking.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)
        st.download_button("⬇ Milestone Report",to_xl({'DQ-Milestone':pv_ms(a_full),'P44 RCA':pv_rca(a_full),'MS Carrier':pv_mc(a_full),'Lane':pv_lane(a_full),'Detail':a_full}),"DQ-Milestone.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True)

    # Apply filters
    a=a_full.copy()
    if f_tracked!="All":a=flt(a,'Tracked',f_tracked)
    if f_status!="All":a=flt(a,'Tracked Status',f_status,True)
    if f_carrier!="All":a=flt(a,'Carrier Name',f_carrier,True)
    if f_lane!="All":a=flt(a,'Lanes',f_lane,True)
    qf=q.copy()
    if f_tracked!="All":qf=flt(qf,'Tracked',f_tracked)
    if f_carrier!="All":qf=flt(qf,'Carrier Name',f_carrier,True)

    # ── KPIs ────────────────────────────────────────────────
    tot=len(a)
    if tot==0:st.warning("No data matches filters.");return
    tl=[su(v) for v in a['Tracked'].tolist()];sl=[ss(v) for v in a['Tracked Status'].tolist()]
    tt=sum(1 for v in tl if v=='TRUE');ft=sum(1 for v in sl if v=='Full Tracked')
    pt=sum(1 for v in sl if v=='Partial Tracked');zm=sum(1 for v in sl if v=='Tracked with 0 milestones')
    tp=tt/tot*100;fp=ft/tot*100;untrk=tot-tt

    st.markdown(f"""<div class="kpi-row">
    <div class="kpi k-blue"><div class="kv">{tot:,}</div><div class="kl">Total Shipments</div></div>
    <div class="kpi k-green"><div class="kv">{tp:.1f}%</div><div class="kl">Tracked</div><div class="kd">{tt:,} shipments</div></div>
    <div class="kpi k-emerald"><div class="kv">{fp:.1f}%</div><div class="kl">Full Track Rate</div><div class="kd">Tracked with all milestones</div></div>
    <div class="kpi k-orange"><div class="kv">{pt:,}</div><div class="kl">Partial Tracked</div><div class="kd">Missing 1+ milestones</div></div>
    <div class="kpi k-red"><div class="kv">{zm:,}</div><div class="kl">Untracked</div><div class="kd">0 milestones received</div></div>
    <div class="kpi k-purple"><div class="kv">{ft:,}</div><div class="kl">Fully Tracked</div><div class="kd">All 4 milestones</div></div>
    </div>""",unsafe_allow_html=True)

    tab1,tab2,tab3=st.tabs(["📡 Tracking Data","🎯 Milestone Data","📋 Processed Data"])

    # ═══════════ TAB 1 ═══════════════════════════════════════
    with tab1:
        c1,c2=st.columns(2)
        with c1:
            st.markdown('<div class="card"><div class="card-h"><h3>Tracking Rate</h3></div><div class="card-b" style="padding:10px 16px">',unsafe_allow_html=True)
            st.plotly_chart(ch_track(a),use_container_width=True,config={'displayModeBar':False})
            st.markdown('</div></div>',unsafe_allow_html=True)
        with c2:
            st.markdown('<div class="card"><div class="card-h"><h3>Top Carriers by Status</h3></div><div class="card-b" style="padding:10px 16px">',unsafe_allow_html=True)
            st.plotly_chart(ch_carrier(a),use_container_width=True,config={'displayModeBar':False})
            st.markdown('</div></div>',unsafe_allow_html=True)

        c3,c4=st.columns(2)
        with c3:
            st.markdown('<div class="card"><div class="card-h"><h3>DQ-Tracking Summary</h3></div><div class="card-b">',unsafe_allow_html=True)
            ts=pv_track(a);td=ts.copy();cm={'TRUE':'#00875a','FALSE':'#de350b'}
            td['Tracked']=[b_tr(v) if v!='Grand Total' else f'<b>{v}</b>' for v in ts['Tracked'].tolist()]
            td['Shipment %']=[pbar(ts.iloc[i]['Shipments']/tot*100,cm.get(ss(ts.iloc[i]['Tracked']),'#666')) if ss(ts.iloc[i]['Tracked'])!='Grand Total' else '<b>100%</b>' for i in range(len(ts))]
            st.markdown(ht(td,180),unsafe_allow_html=True)
            st.markdown('</div></div>',unsafe_allow_html=True)
        with c4:
            st.markdown('<div class="card"><div class="card-h"><h3>Tracking by Carrier</h3></div><div class="card-b">',unsafe_allow_html=True)
            tc=pv_track_carr(a);tcd=tc.copy();tcd['Tracked']=[b_tr(v) for v in tc['Tracked'].tolist()]
            st.markdown(ht(tcd,220),unsafe_allow_html=True)
            st.markdown('</div></div>',unsafe_allow_html=True)

        st.markdown(f'<div class="card"><div class="card-h"><h3>Tracking Detail</h3><span class="cnt">{len(a):,} shipments</span></div><div class="card-b">',unsafe_allow_html=True)
        cols=['Carrier Name','Bill of Lading','Tracked','Connection Type','Tracking Method','Pickup Location','Destination Location','Lanes','Final Status Reason','Tracked Status','Tracking Error']
        dd=a[cols].copy();dd['Tracked']=[b_tr(v) for v in dd['Tracked'].tolist()];dd['Tracked Status']=[b_st(v) for v in dd['Tracked Status'].tolist()]
        st.markdown(ht(dd,420),unsafe_allow_html=True)
        st.markdown('</div></div>',unsafe_allow_html=True)
        st.download_button("⬇ Export Filtered Tracking Data",to_xl({'Tracking Detail':a}),"Tracking-Filtered.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",key="dl_t")

    # ═══════════ TAB 2 ═══════════════════════════════════════
    with tab2:
        c1,c2=st.columns(2)
        with c1:
            st.markdown('<div class="card"><div class="card-h"><h3>Milestone Completeness</h3></div><div class="card-b" style="padding:10px 16px">',unsafe_allow_html=True)
            st.plotly_chart(ch_ms(a),use_container_width=True,config={'displayModeBar':False})
            st.markdown('</div></div>',unsafe_allow_html=True)
        with c2:
            st.markdown('<div class="card"><div class="card-h"><h3>P44 Root Cause Analysis</h3></div><div class="card-b" style="padding:10px 16px">',unsafe_allow_html=True)
            st.plotly_chart(ch_rca(a),use_container_width=True,config={'displayModeBar':False})
            st.markdown('</div></div>',unsafe_allow_html=True)

        c3,c4=st.columns(2)
        with c3:
            st.markdown('<div class="card"><div class="card-h"><h3>DQ-Milestone Summary</h3></div><div class="card-b">',unsafe_allow_html=True)
            ms=pv_ms(a);md=ms.copy();cm2={'Full Tracked':'#00875a','Partial Tracked':'#ff8b00','Tracked with 0 milestones':'#de350b'}
            md['Tracked Status']=[b_st(v) if v!='Grand Total' else f'<b>{v}</b>' for v in ms['Tracked Status'].tolist()]
            md['Shipment %']=[pbar(ms.iloc[i]['Shipments']/tot*100,cm2.get(ss(ms.iloc[i]['Tracked Status']),'#666')) if ss(ms.iloc[i]['Tracked Status'])!='Grand Total' else '<b>100%</b>' for i in range(len(ms))]
            st.markdown(ht(md,180),unsafe_allow_html=True)
            st.markdown('</div></div>',unsafe_allow_html=True)
        with c4:
            st.markdown('<div class="card"><div class="card-h"><h3>P44 Root Cause Analysis</h3></div><div class="card-b">',unsafe_allow_html=True)
            rc=pv_rca(a);rd=rc.copy();rd['P44 Analysis']=[b_p4(v) if ss(v)!='Grand Total' else f'<b>{v}</b>' for v in rc['P44 Analysis'].tolist()]
            st.markdown(ht(rd,220),unsafe_allow_html=True)
            st.markdown('</div></div>',unsafe_allow_html=True)

        c5,c6=st.columns(2)
        with c5:
            st.markdown('<div class="card"><div class="card-h"><h3>Milestone by Carrier</h3></div><div class="card-b">',unsafe_allow_html=True)
            mc=pv_mc(a);mcd=mc.copy();mcd['Milestone Missed']=[b_mm(v) for v in mc['Milestone Missed'].tolist()]
            st.markdown(ht(mcd,220),unsafe_allow_html=True)
            st.markdown('</div></div>',unsafe_allow_html=True)
        with c6:
            st.markdown('<div class="card"><div class="card-h"><h3>Lane Analysis</h3></div><div class="card-b">',unsafe_allow_html=True)
            la=pv_lane(a);lad=la.copy();lad['Milestone Missed']=[b_mm(v) for v in la['Milestone Missed'].tolist()]
            st.markdown(ht(lad,220),unsafe_allow_html=True)
            st.markdown('</div></div>',unsafe_allow_html=True)

        mcols=['Carrier Name','Bill of Lading','Tracked','Lanes','Pickup Location','Destination Location','Pickup Arrival Milestone (UTC)','Pickup Departure Milestone (UTC)','Final Destination Arrival Milestone (UTC)','Final Destination Departure Milestone (UTC)','Final Status Reason','Tracked Status','Milestone Achieved','Milestone Missed','p44 Analysis']
        st.markdown(f'<div class="card"><div class="card-h"><h3>Milestone Detail</h3><span class="cnt">{len(a):,} shipments</span></div><div class="card-b">',unsafe_allow_html=True)
        mdd=a[mcols].copy();mdd['Tracked']=[b_tr(v) for v in mdd['Tracked'].tolist()];mdd['Tracked Status']=[b_st(v) for v in mdd['Tracked Status'].tolist()]
        mdd['Milestone Missed']=[b_mm(v) for v in mdd['Milestone Missed'].tolist()];mdd['p44 Analysis']=[b_p4(v) for v in mdd['p44 Analysis'].tolist()]
        st.markdown(ht(mdd,420),unsafe_allow_html=True)
        st.markdown('</div></div>',unsafe_allow_html=True)
        st.download_button("⬇ Export Filtered Milestone Data",to_xl({'Milestone Detail':a}),"Milestone-Filtered.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",key="dl_m")

    # ═══════════ TAB 3 ═══════════════════════════════════════
    with tab3:
        st.markdown(f'<div class="card"><div class="card-h"><h3>Full Processed Data (Query Sheet)</h3><span class="cnt">{len(qf):,} shipments</span></div><div class="card-b">',unsafe_allow_html=True)
        rdd=qf.copy();rdd['Tracked']=[b_tr(v) for v in rdd['Tracked'].tolist()]
        st.markdown(ht(rdd,520),unsafe_allow_html=True)
        st.markdown('</div></div>',unsafe_allow_html=True)
        st.download_button("⬇ Export Filtered Data",to_xl({'Processed Data':qf}),"Processed-Filtered.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",key="dl_r")

if __name__=='__main__':
    main()

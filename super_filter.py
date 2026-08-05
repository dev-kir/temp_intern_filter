import argparse, json, math, statistics
from pathlib import Path
from collections import Counter
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.utils import get_column_letter

# ═══════════════════════════════════════════════════════════════════════════
# CONFIG (hardcoded — edit INPUT_FILE / OUTPUT_FILE here)
# ═══════════════════════════════════════════════════════════════════════════

INPUT_FILE = "SARI_Results_2026-08-04-00-36-28.xlsx"
OUTPUT_FILE = "SARI_Organisation.xlsx"

CFG = {
    "high_consensus": 0.80,
    "moderate_consensus": 0.60,
    "minimum_respondents_for_agreement": 2,
    "tier_boundaries": [0.20, 0.40, 0.60, 0.80],
    "section_order": [
        "Strategy & Leadership",
        "Governance, Policy & Ethics",
        "Talent & Organisational Culture",
        "Infrastructure & Technology",
        "Data Management & Readiness",
        "AI Implementation & Potential Impact",
        "Investment",
    ],
    "section_prefixes": {
        "strategy": "Strategy & Leadership",
        "governance": "Governance, Policy & Ethics",
        "talent": "Talent & Organisational Culture",
        "infrastructure": "Infrastructure & Technology",
        "data": "Data Management & Readiness",
        "aiapp": "AI Implementation & Potential Impact",
        "investment": "Investment",
        "background": "Background",
    },
    "multi_select_question_ids": ["background_2", "background_3", "background_4"],
}
NAVY='FF17365D'; BLUE='FF2F75B5'; WHITE='FFFFFFFF'; LIGHT='FFDDEBF7'; GREEN='FFE2F0D9'; GRAY='FFF2F2F2'; RED='FFF4CCCC'
REQ=['Respondent ID','Submitted at','Question #','Question ID','Question','Answer','Answer value','Answer score','Max score','Participant name','Job title','Organisation name','Organisation type','Organisation size','Stakeholder category','PCDS sector','District','Role level','Department','Age band','Part of group','Parent company']
RAW_OUT=['Respondent ID','Submitted at','Standard section','Question #','Question ID','Question','Answer','Answer value','Answer score','Max score','Participant name','Job title','Organisation name','Organisation type','Organisation size','Stakeholder category','PCDS sector','District','Role level','Department','Age band','Part of group','Parent company']

def tier(x,b):
    if x < b[0]: return 'AI Aware - 0'
    if x < b[1]: return 'AI Explorer - 1'
    if x < b[2]: return 'AI Follower - 2'
    if x < b[3]: return 'AI Leader - 3'
    return 'AI Pioneer - 4'

def agreement(consensus,n,min_n,high,moderate):
    if n < min_n: return 'Not measurable'
    if consensus >= high: return 'High'
    if consensus >= moderate: return 'Moderate'
    return 'Low'

def mode_det(values):
    vals=[str(v).strip() for v in values if pd.notna(v) and str(v).strip()]
    if not vals: return '',0
    c=Counter(vals); m=max(c.values()); winners=sorted(k for k,v in c.items() if v==m)
    return winners[0],m

def title(ws,text,subtitle,span,font_size=16,font_name='Calibri',row1_height=27.75,row2_height=30):
    ws.sheet_view.showGridLines=False
    ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=span)
    if text is not None:
        ws['A1']=text; ws['A1'].font=Font(name=font_name,size=font_size,bold=True,color=WHITE); ws['A1'].fill=PatternFill('solid',fgColor=NAVY); ws['A1'].alignment=Alignment(vertical='center')
    ws.row_dimensions[1].height=row1_height
    ws.merge_cells(start_row=2,start_column=1,end_row=2,end_column=span)
    ws['A2']=subtitle; ws['A2'].font=Font(size=10,color='666666'); ws['A2'].alignment=Alignment(wrap_text=True); ws.row_dimensions[2].height=row2_height

def headers(ws,row,n):
    for c in range(1,n+1):
        x=ws.cell(row,c); x.fill=PatternFill('solid',fgColor=BLUE); x.font=Font(bold=True,color=WHITE); x.alignment=Alignment(wrap_text=True,vertical='center')
    ws.row_dimensions[row].height=32

def write_table(ws,columns,rows,widths=None,freeze='A5',add_filter=True):
    for c,h in enumerate(columns,1): ws.cell(4,c,h)
    headers(ws,4,len(columns))
    for r,row in enumerate(rows,5):
        for c,v in enumerate(row,1): ws.cell(r,c,v)
    ws.freeze_panes=freeze
    if add_filter:
        ws.auto_filter.ref=f'A4:{get_column_letter(len(columns))}{4+len(rows)}'
    if widths:
        for c,w in enumerate(widths,1): ws.column_dimensions[get_column_letter(c)].width=w

def load_data(path,cfg):
    df=pd.read_excel(path,sheet_name='Answers',engine='openpyxl',dtype=object)
    missing=[c for c in REQ if c not in df.columns]
    if missing: raise ValueError('Missing required columns: '+', '.join(missing))
    for c in ['Answer score','Max score','Question #']: df[c]=pd.to_numeric(df[c],errors='coerce')
    df=df[df['Organisation name'].notna()].copy()
    df['Organisation name']=df['Organisation name'].astype(str).str.strip()
    df=df[df['Organisation name']!='']
    df=df.drop_duplicates(['Respondent ID','Question ID'],keep='last')
    pref=df['Question ID'].astype(str).str.split('_').str[0]
    df['Standard section']=pref.map(cfg['section_prefixes']).fillna(df.get('Section','Other'))
    return df

def calculate(df,cfg):
    high=cfg['high_consensus']; mod=cfg['moderate_consensus']; min_n=cfg['minimum_respondents_for_agreement']; sections=cfg['section_order']; multi=set(cfg['multi_select_question_ids'])
    qrows=[]; qlookup={}
    for (org,sec,qid,qtext),g in df.groupby(['Organisation name','Standard section','Question ID','Question'],dropna=False,sort=True):
        n=g['Respondent ID'].nunique(); mode,count=mode_det(g['Answer']); cons=count/n if n else 0
        scored=(g['Max score'].fillna(0)>0).any(); scores=g.loc[g['Max score'].fillna(0)>0,'Answer score'].dropna().astype(float)
        maxs=g.loc[g['Max score'].fillna(0)>0,'Max score'].dropna().astype(float)
        avg=scores.mean() if len(scores) else None; med=scores.median() if len(scores) else None; mn=scores.min() if len(scores) else None; mx=scores.max() if len(scores) else None
        rng=(mx-mn) if len(scores) else None; sd=scores.std(ddof=1) if len(scores)>1 else (0 if len(scores)==1 else None); norm=(avg/maxs.mean()) if len(scores) and maxs.mean() else None
        agr=agreement(cons,n,min_n,high,mod); flag='Review' if n>=min_n and cons<mod else ''
        row=[org,sec,qid,qtext,n,bool(scored),mode,count,cons,avg,med,mn,mx,rng,sd,norm,agr,flag]
        qrows.append(row); qlookup[(org,sec,qid)]=row
    qcols=['Organisation name','Section','Question ID','Question','Respondents','Scored question','Most common answer','Most common count','Consensus','Average score','Median score','Minimum score','Maximum score','Score range','Standard deviation','Normalised score','Agreement','Review flag']
    qdf=pd.DataFrame(qrows,columns=qcols)

    srows=[]
    for org in sorted(df['Organisation name'].unique()):
        for sec in sections:
            g=df[(df['Organisation name']==org)&(df['Standard section']==sec)&(df['Max score'].fillna(0)>0)]
            if g.empty: continue
            scores=g['Answer score'].dropna().astype(float); maxs=g['Max score'].dropna().astype(float); n=g['Respondent ID'].nunique(); nq=g['Question ID'].nunique()
            avg=scores.mean(); med=scores.median(); mn=scores.min(); mx=scores.max(); maxp=maxs.mean(); norm=avg/maxp if maxp else None
            qs=qdf[(qdf['Organisation name']==org)&(qdf['Section']==sec)]; cons=qs['Consensus'].mean() if len(qs) else None
            agr=agreement(cons or 0,n,min_n,high,mod)
            srows.append([org,sec,n,nq,avg,med,mn,mx,maxp,norm,cons,agr])
    scols=['Organisation name','Section','Respondents','Questions','Average score','Median score','Minimum score','Maximum score','Max possible','Normalised score','Average consensus','Agreement']
    sdf=pd.DataFrame(srows,columns=scols)

    orows=[]
    for org,g in df.groupby('Organisation name',sort=True):
        n=g['Respondent ID'].nunique(); scored=g[g['Max score'].fillna(0)>0]; avg=scored['Answer score'].dropna().astype(float).mean(); denom=scored['Max score'].dropna().astype(float).mean(); overall=avg/denom if denom else 0
        ss=sdf[sdf['Organisation name']==org]; strongest=ss.sort_values(['Normalised score','Section'],ascending=[False,True]).iloc[0]['Section'] if len(ss) else ''; weakest=ss.sort_values(['Normalised score','Section']).iloc[0]['Section'] if len(ss) else ''
        qs=qdf[qdf['Organisation name']==org]; cons=qs['Consensus'].mean() if len(qs) else 0; reviews=int((qs['Review flag']=='Review').sum()); agr=agreement(cons,n,min_n,high,mod)
        first=g.iloc[0]
        typ=first.get('Organisation type',''); size=first.get('Organisation size',''); sector=first.get('PCDS sector',''); latest=g.iloc[-1].get('Submitted at','')
        depts=g['Department'].dropna().astype(str).replace('',pd.NA).dropna().nunique(); roles=g['Role level'].dropna().astype(str).replace('',pd.NA).dropna().nunique()
        interpretation='Single respondent: perception only' if n<2 else ('Directional: small sample' if n<3 else 'Multi-respondent view')
        b=cfg['tier_boundaries']; nxt=(1-overall) if overall>=b[-1] else min(x for x in b if x>overall)-overall
        orows.append([org,typ,n,depts,roles,size,sector,latest,avg,overall,strongest,weakest,cons,reviews,agr,interpretation,tier(overall,b),nxt])
    ocols=['Organisation name','Organisation type','Respondents','Departments represented','Role levels represented','Organisation size','Sector','Latest submission','Average score','Overall score','Strongest section','Weakest section','Average consensus','Questions for review','Agreement','Interpretation','Maturity tier','Distance to next tier']
    odf=pd.DataFrame(orows,columns=ocols)

    drows=[]
    for (org,sec,qid,qtext),g in df.groupby(['Organisation name','Standard section','Question ID','Question'],sort=True):
        n=g['Respondent ID'].nunique()
        if qid in multi:
            pairs=[]
            for _,x in g.iterrows():
                for opt in str(x['Answer'] or '').split(','):
                    if opt.strip(): pairs.append((x['Respondent ID'],opt.strip(),x['Answer score']))
            for opt in sorted(set(p[1] for p in pairs)):
                selected=len(set(p[0] for p in pairs if p[1]==opt)); scores=[p[2] for p in pairs if p[1]==opt and pd.notna(p[2])]
                drows.append([org,sec,qid,qtext,'Multi-select',opt,(sum(scores)/len(scores) if scores else None),selected,selected/n if n else 0,n])
        else:
            for opt,gg in g.groupby('Answer',dropna=False,sort=True):
                selected=gg['Respondent ID'].nunique(); sc=gg['Answer score'].dropna().astype(float)
                drows.append([org,sec,qid,qtext,'Single-choice',opt,(sc.mean() if len(sc) else None),selected,selected/n if n else 0,n])
    dcols=['Organisation name','Standard section','Question ID','Question','Question type','Answer option','Answer score','Respondents selecting','Percentage','Question respondents']
    ddf=pd.DataFrame(drows,columns=dcols)
    return odf,sdf,qdf,ddf

def build(df,odf,sdf,qdf,ddf,cfg,out):
    wb=Workbook(); wb.remove(wb.active); sections=cfg['section_order']
    # Read Me
    ws=wb.create_sheet('Read Me'); title(ws,'SARI Organisation Statistics','Interactive organisation-level reporting workbook generated from the uploaded SARI results.',8,font_size=16,font_name='Calibri',row1_height=27.75)
    notes=[
        ('Purpose','Analyse questionnaire results by organisation, section, question and answer distribution.'),
        ('Dashboard','Choose an organisation in cell B4. The section score table and chart update using formulas.'),
        ('Organisation Summary','One row per organisation, including respondent count, maturity score, strongest/weakest sections and agreement.'),
        ('Section Summary','Scored question results aggregated by organisation and standardised section.'),
        ('Question Summary','Question-level mode, consensus, score statistics, agreement and review flags.'),
        ('Answer Distribution','Counts and percentages for each answer option. Multi-select percentages may exceed 100% in total.'),
        ('Scoring rule','Only rows where Max score is greater than zero are included in maturity scores. Background questions remain distribution-only.'),
        ('Agreement rule','High: at least 80% selected the modal answer; Moderate: 60% to below 80%; Low: below 60%; not measurable for one respondent.'),
        ('Caution','Results with one respondent represent an individual perception, not organisation-wide consensus.'),
        ('Data standardisation','Sections are standardised using the Question ID prefix, preventing English and Malay labels from splitting the same section.'),
        ('Refresh','This workbook is a snapshot. Re-run the analysis when new responses are added to the source export.'),
        ('Enhancements in this version',''),
        ('','Maturity tiers, distance to next tier and a printable Organisation Report are included.'),
    ]
    for r,(a,b) in enumerate(notes,4): ws.cell(r,1,a).font=Font(bold=True,color=NAVY); ws.cell(r,2,b); ws.merge_cells(start_row=r,start_column=2,end_row=r,end_column=8)
    ws.column_dimensions['A'].width=24; ws.column_dimensions['B'].width=100
    ws.page_setup.orientation='portrait'
    # Lists
    ws=wb.create_sheet('Lists');
    for r,o in enumerate(odf['Organisation name'],1): ws.cell(r,1,o)
    ws.sheet_state='hidden'
    ws.page_setup.orientation='portrait'
    # Dashboard
    ws=wb.create_sheet('Dashboard'); title(ws,'Organisation Dashboard','Select an organisation to view the current results.',10,font_size=16,font_name='Calibri',row1_height=27.75)
    ws['A4']='Selected organisation'; ws['B4']=odf.iloc[0]['Organisation name'] if len(odf) else ''; ws['B4'].fill=PatternFill('solid',fgColor=GREEN); ws['B4'].font=Font(bold=True,color='008000')
    items=[('A6','Respondents','A7','C'),('D6','Overall score','D7','J'),('G6','Strongest section','G7','K'),('A9','Weakest section','A10','L'),('D9','Agreement','D10','O'),('G9','Interpretation','G10','P'),('A12','Maturity tier','A13','Q'),('D12','Distance to next tier','D13','R'),('G12','Questions for review','G13','N')]
    for lc,label,vc,col in items:
        ws[lc]=label; ws[lc].font=Font(bold=True,color=NAVY); ws[vc]=f'=IFERROR(INDEX(\'Organisation Summary\'!${col}:${col},MATCH($B$4,\'Organisation Summary\'!$A:$A,0)),"")'; ws[vc].fill=PatternFill('solid',fgColor=LIGHT); ws[vc].font=Font(size=13,bold=True,color=NAVY); ws[vc].alignment=Alignment(wrap_text=True)
    ws['D7'].number_format=ws['D13'].number_format='0.0%'
    hs=['Section','Average score','Normalised score','Respondents','Agreement']
    for c,h in enumerate(hs,1): ws.cell(17,c,h)
    headers(ws,17,5)
    for r,s in enumerate(sections,18):
        ws.cell(r,1,s); ws.cell(r,2,f'=IFERROR(SUMIFS(\'Section Summary\'!$E:$E,\'Section Summary\'!$A:$A,$B$4,\'Section Summary\'!$B:$B,$A{r}),"")'); ws.cell(r,3,f'=IFERROR(SUMIFS(\'Section Summary\'!$J:$J,\'Section Summary\'!$A:$A,$B$4,\'Section Summary\'!$B:$B,$A{r}),"")'); ws.cell(r,4,f'=IFERROR(SUMIFS(\'Section Summary\'!$C:$C,\'Section Summary\'!$A:$A,$B$4,\'Section Summary\'!$B:$B,$A{r}),"")'); ws.cell(r,5,f'=IFERROR(LOOKUP(2,1/((\'Section Summary\'!$A$5:$A${4+len(sdf)}=$B$4)*(\'Section Summary\'!$B$5:$B${4+len(sdf)}=$A{r})),\'Section Summary\'!$L$5:$L${4+len(sdf)}),"")'); ws.cell(r,3).number_format='0.0%'
    ch=BarChart(); ch.type='bar'; ch.title='Section maturity profile'; ch.add_data(Reference(ws,min_col=3,min_row=17,max_row=24),titles_from_data=True); ch.set_categories(Reference(ws,min_col=1,min_row=18,max_row=24)); ch.height=7.5; ch.width=15; ch.legend=None; ws.add_chart(ch,'G17'); ws.freeze_panes='A17'
    for c,w in {'A':42,'B':39.140625,'C':14,'D':15,'E':18,'F':13,'G':31.42578125,'H':13,'I':13,'J':13}.items(): ws.column_dimensions[c].width=w
    ws.page_setup.orientation='portrait'
    # Organisation Report
    ws=wb.create_sheet('Organisation Report'); title(ws,'SARI Organisation Report','Select an organisation. The page is formatted for printing.',8,font_size=20,font_name='Cambria',row1_height=33.75)
    ws['A3']='Selected organisation'; ws['B3']=odf.iloc[0]['Organisation name'] if len(odf) else ''; ws['B3'].fill=PatternFill('solid',fgColor=GREEN); ws['B3'].font=Font(bold=True,color='008000')
    items=[('A5','Overall score','B5','J'),('D5','Maturity tier','E5','Q'),('G5','Respondents','H5','C'),('A7','Strongest section','B7','K'),('D7','Weakest section','E7','L'),('G7','Agreement','H7','O')]
    for lc,label,vc,col in items: ws[lc]=label; ws[lc].font=Font(bold=True,color=NAVY); ws[vc]=f'=IFERROR(INDEX(\'Organisation Summary\'!${col}:${col},MATCH($B$3,\'Organisation Summary\'!$A:$A,0)),"")'; ws[vc].fill=PatternFill('solid',fgColor=LIGHT); ws[vc].alignment=Alignment(wrap_text=True)
    ws['B5'].number_format='0.0%';
    for c,h in enumerate(['Section','Score','Agreement'],1): ws.cell(10,c,h)
    headers(ws,10,3)
    for r,s in enumerate(sections,11): ws.cell(r,1,s); ws.cell(r,2,f'=IFERROR(SUMIFS(\'Section Summary\'!$J:$J,\'Section Summary\'!$A:$A,$B$3,\'Section Summary\'!$B:$B,$A{r}),"")'); ws.cell(r,3,f'=IFERROR(LOOKUP(2,1/((\'Section Summary\'!$A$5:$A${4+len(sdf)}=$B$3)*(\'Section Summary\'!$B$5:$B${4+len(sdf)}=$A{r})),\'Section Summary\'!$L$5:$L${4+len(sdf)}),"")'); ws.cell(r,2).number_format='0.0%'
    ws['A20']='Priority questions for review'; ws['A20'].font=Font(size=12,bold=True,color=NAVY)
    for c,h in enumerate(['Question ID','Question','Most common answer','Normalised score','Agreement','Review flag'],1): ws.cell(21,c,h)
    headers(ws,21,6)
    for i,r in enumerate(range(22,27),1):
        for c,col in enumerate(['B','C','D','E','F','G'],1): ws.cell(r,c,f'=IFERROR(INDEX(\'Priority Detail\'!${col}:${col},MATCH($B$3&"|{i}",\'Priority Detail\'!$I:$I,0)),"")')
        ws.cell(r,4).number_format='0.0%'
    for c,w in enumerate([35,28,13,20,18,13,13,13],1): ws.column_dimensions[get_column_letter(c)].width=w
    ws.freeze_panes='A10'
    ws.page_setup.orientation='landscape'; ws.page_setup.fitToWidth=1; ws.page_setup.fitToHeight=1; ws.sheet_properties.pageSetUpPr.fitToPage=True; ws.print_area='A1:H27'
    # Data sheets
    for name,sub,frame,widths,freeze,add_filter in [
        ('Organisation Summary','One row per organisation. Sort or filter by respondent count, score or agreement.',odf,[47.7109375,37,14.140625,25,22.7109375,17.42578125,25.7109375,18.140625,17,14.7109375,27.28515625,13,19.42578125,20.7109375,12.5703125,25.42578125,20,13],'A23',True),
        ('Section Summary','Section-level maturity and internal agreement for each organisation.',sdf,[18.28515625,15.7109375,13,13,13,13,13,13,13,13,13,13],'A5',False),
        ('Question Summary','Question-level statistics. Use Review flag to find low-consensus scored questions.',qdf,[45,38,13,72,12.28515625,13,60,13,10.85546875,13,13,10.85546875,9.7109375,13,13,11.140625,11.5703125,13],'D5',False),
        ('Answer Distribution','Counts and percentages by answer option. Multi-select totals can exceed 100%.',ddf,[45,20.5703125,13.140625,72,17.28515625,65,13.7109375,16.28515625,10.85546875,16],'A5',False),
    ]:
        ws=wb.create_sheet(name); title(ws,name,sub,len(frame.columns),font_size=16,font_name='Calibri',row1_height=27.75); write_table(ws,list(frame.columns),frame.where(pd.notna(frame),None).values.tolist(),widths,freeze,add_filter)
        ws.page_setup.orientation='portrait'
        if name=='Organisation Summary':
            for r in range(5,5+len(frame)): ws.cell(r,10).number_format=ws.cell(r,13).number_format=ws.cell(r,18).number_format='0.0%'
        if name=='Section Summary':
            for r in range(5,5+len(frame)): ws.cell(r,10).number_format=ws.cell(r,11).number_format='0.0%'
        if name=='Question Summary':
            for r in range(5,5+len(frame)): ws.cell(r,9).number_format=ws.cell(r,16).number_format='0.0%'
        if name=='Answer Distribution':
            for r in range(5,5+len(frame)): ws.cell(r,9).number_format='0.0%'
    # Raw answers
    rframe=df.copy();
    for c in RAW_OUT:
        if c not in rframe.columns: rframe[c]=None
    rframe=rframe[RAW_OUT]
    ws=wb.create_sheet('Raw Answers'); title(ws,None,'Imported source rows with an added Standard section field. Personal email is intentionally omitted from this analysis copy.',len(RAW_OUT),font_size=16,font_name='Calibri',row1_height=27.75); write_table(ws,RAW_OUT,rframe.where(pd.notna(rframe),None).values.tolist(),[22.85546875,17.7109375,38,10.7109375,13.28515625,70,65,14.7109375,12,12.7109375,28,13,45,17.140625,15.42578125,20.140625,10.7109375,13,13,12.7109375,13,13,14.5703125],'A5',False)
    # Apply navy fill to A1 even though no title text
    ws['A1'].fill=PatternFill('solid',fgColor=NAVY)
    ws['A1'].font=Font(name='Calibri',size=16,bold=True,color=WHITE)
    ws['A1'].alignment=Alignment(vertical='center')
    ws.page_setup.orientation='portrait'
    # Priority helper
    rows=[]
    for org,g in qdf[qdf['Scored question']==True].groupby('Organisation name'):
        gg=g.copy(); gg['_review']=(gg['Review flag']=='Review').astype(int); gg=gg.sort_values(['_review','Normalised score','Question ID'],ascending=[False,True,True])
        for rank,(_,x) in enumerate(gg.iterrows(),1): rows.append([org,x['Question ID'],x['Question'],x['Most common answer'],x['Normalised score'],x['Agreement'],x['Review flag'],rank,f'{org}|{rank}'])
    ws=wb.create_sheet('Priority Detail'); cols=['Organisation','Question ID','Question','Most common answer','Normalised score','Agreement','Review flag','Priority rank','Lookup key'];
    for c,h in enumerate(cols,1): ws.cell(1,c,h)
    for r,row in enumerate(rows,2):
        for c,v in enumerate(row,1): ws.cell(r,c,v)
    ws.sheet_state='hidden'
    ws.page_setup.orientation='portrait'
    wb.calculation.fullCalcOnLoad=True; wb.calculation.forceFullCalc=True; wb.calculation.calcMode='auto'; wb.active=wb.sheetnames.index('Dashboard'); wb.save(out)

def main():
    print(f"Reading: {INPUT_FILE}")
    df = load_data(INPUT_FILE, CFG)
    print(f"  {len(df)} rows after dedup, {df['Respondent ID'].nunique()} respondents")
    print("Calculating...")
    odf, sdf, qdf, ddf = calculate(df, CFG)
    print(f"  {len(odf)} organisations, {len(sdf)} section rows, {len(qdf)} question rows")
    print("Building workbook...")
    build(df, odf, sdf, qdf, ddf, CFG, OUTPUT_FILE)
    print(f"Saved: {OUTPUT_FILE}")
if __name__=='__main__': main()

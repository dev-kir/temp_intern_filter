#!/usr/bin/env python3
"""Generate the complete SARI_Organisation workbook from a raw SARI export.

Architecture: the supplied report_template.xlsx controls visual formatting, charts,
page setup and dashboard layout. This script rebuilds every data/output sheet from
Answers and Scores, repairs the organisation dropdowns, and writes a new workbook.
"""
from __future__ import annotations
import argparse, math, shutil
from pathlib import Path
from collections import Counter, defaultdict
from copy import copy
from statistics import mean, median, stdev
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName

SECTIONS={
 'strategy':'Strategy & Leadership','governance':'Governance, Policy & Ethics',
 'talent':'Talent & Organisational Culture','infrastructure':'Infrastructure & Technology',
 'data':'Data Management & Readiness','aiapp':'AI Implementation & Potential Impact',
 'investment':'Investment','background':'Background'}
SCORED=[v for k,v in SECTIONS.items() if k!='background']
HIGH=.8; MOD=.6; MIN_RESP=2

def norm_text(x): return '' if x is None else str(x).strip()
def num(x): return isinstance(x,(int,float)) and not isinstance(x,bool)
def standard_section(qid, fallback=''):
    prefix=norm_text(qid).split('_')[0].lower()
    return SECTIONS.get(prefix, norm_text(fallback) or 'Other')
def agreement(n,cons):
    if n<MIN_RESP:return 'Not measurable'
    return 'High' if cons>=HIGH else 'Moderate' if cons>=MOD else 'Low'
def maturity(x):
    return 'AI Aware - 0' if x<.2 else 'AI Explorer - 1' if x<.4 else 'AI Follower - 2' if x<.6 else 'AI Leader - 3' if x<.8 else 'AI Pioneer - 4'
def distance(x): return 1-x if x>=.8 else math.ceil((x+1e-12)/.2)*.2-x
def safe_stdev(v): return stdev(v) if len(v)>1 else 0

def copy_row_style(ws,src,dst,max_col):
    for c in range(1,max_col+1):
        ws.cell(dst,c)._style=copy(ws.cell(src,c)._style)
        ws.cell(dst,c).number_format=ws.cell(src,c).number_format
        ws.cell(dst,c).alignment=copy(ws.cell(src,c).alignment)
    ws.row_dimensions[dst].height=ws.row_dimensions[src].height

def rewrite(ws, headers, rows, start=5):
    max_col=len(headers)
    # preserve row 5 as style source before clearing
    for r in range(start, max(ws.max_row,start)+1):
        for c in range(1,max(ws.max_column,max_col)+1): ws.cell(r,c).value=None
    for c,h in enumerate(headers,1): ws.cell(start-1,c,h)
    for i,row in enumerate(rows,start):
        if i!=start: copy_row_style(ws,start,i,max_col)
        for c,v in enumerate(row,1): ws.cell(i,c,v)
    ws.auto_filter.ref=f'A{start-1}:{ws.cell(start-1,max_col).column_letter}{start+len(rows)-1}'

def read_source(path):
    wb=load_workbook(path,read_only=True,data_only=True)
    if 'Answers' not in wb.sheetnames or 'Scores' not in wb.sheetnames:
        raise ValueError('Input must contain Answers and Scores sheets')
    def read(ws):
        h=[c.value for c in ws[1]]; return h,[dict(zip(h,r)) for r in ws.iter_rows(min_row=2,values_only=True)]
    ah,answers=read(wb['Answers']); sh,scores=read(wb['Scores'])
    # discard empty records and deduplicate respondent-question, keeping last occurrence
    dedup={}
    for a in answers:
        if norm_text(a.get('Respondent ID')) and norm_text(a.get('Question ID')):
            dedup[(norm_text(a['Respondent ID']),norm_text(a['Question ID']))]=a
    answers=list(dedup.values())
    return ah,answers,sh,[s for s in scores if norm_text(s.get('Respondent ID'))]

def build(input_file,template_file,output_file):
    ah,answers,sh,scores=read_source(input_file)
    shutil.copy2(template_file,output_file)
    wb=load_workbook(output_file,data_only=False)
    # Ensure this is the intended clean workbook structure
    for bad in ['Organisation Type','Section Benchmark']:
        if bad in wb.sheetnames: del wb[bad]
    by_org_answers=defaultdict(list); by_org_scores=defaultdict(list)
    for a in answers:
        org=norm_text(a.get('Organisation name'))
        if org:
            a['_section']=standard_section(a.get('Question ID'),a.get('Section')); by_org_answers[org].append(a)
    for s in scores:
        org=norm_text(s.get('Organisation name'))
        if org: by_org_scores[org].append(s)
    orgs=sorted(set(by_org_answers)|set(by_org_scores),key=str.casefold)

    # Raw Answers, excluding personal email, with Standard section inserted
    raw_headers=['Respondent ID','Submitted at','Standard section','Question #','Question ID','Question','Answer','Answer value','Answer score','Max score','Participant name','Job title','Organisation name','Organisation type','Organisation size','Stakeholder category','PCDS sector','District','Role level','Department','Age band','Part of group','Parent company']
    raw_rows=[]
    for a in answers:
        raw_rows.append([a.get(h) if h!='Standard section' else a['_section'] for h in raw_headers])
    raw_rows.sort(key=lambda r:(norm_text(r[12]).casefold(),norm_text(r[0]),norm_text(r[4])))
    rewrite(wb['Raw Answers'],raw_headers,raw_rows)

    # Question summary and answer distribution
    qgroups=defaultdict(list)
    for a in answers:
        org=norm_text(a.get('Organisation name')); qid=norm_text(a.get('Question ID'))
        if org and qid:qgroups[(org,a['_section'],qid,norm_text(a.get('Question'))) ].append(a)
    qrows=[]; distrows=[]; q_lookup={}
    for key,grp in sorted(qgroups.items(),key=lambda x:(x[0][0].casefold(),x[0][1],x[0][2])):
        org,sec,qid,qtext=key
        respondents=len({norm_text(x.get('Respondent ID')) for x in grp})
        answers_list=[norm_text(x.get('Answer')) for x in grp if norm_text(x.get('Answer'))]
        counts=Counter(answers_list)
        mode,mode_count=('',0) if not counts else max(counts.items(),key=lambda kv:(kv[1],-answers_list.index(kv[0])))
        consensus=mode_count/respondents if respondents else 0
        scored=any(num(x.get('Max score')) and x.get('Max score')>0 for x in grp)
        vals=[float(x.get('Answer score')) for x in grp if scored and num(x.get('Answer score'))]
        mx=max([float(x.get('Max score')) for x in grp if num(x.get('Max score'))] or [0])
        av=mean(vals) if vals else None; med=median(vals) if vals else None; mn=min(vals) if vals else None; ma=max(vals) if vals else None
        rng=ma-mn if vals else None; sd=safe_stdev(vals) if vals else None; ns=av/mx if vals and mx else None
        agr=agreement(respondents,consensus); review='Review' if respondents>=MIN_RESP and consensus<MOD else ''
        row=[org,sec,qid,qtext,respondents,scored,mode,mode_count,consensus,av,med,mn,ma,rng,sd,ns,agr,review]
        qrows.append(row); q_lookup[(org,sec,qid)]=row
        # Full answer distribution. Multi-select options are split by commas.
        multi=qid in {'background_2','background_3','background_4'}
        option_counts=Counter()
        score_by_option=defaultdict(list)
        for x in grp:
            ans=norm_text(x.get('Answer'))
            opts=[z.strip() for z in ans.split(',')] if multi else [ans]
            for opt in filter(None,opts):
                option_counts[opt]+=1
                if num(x.get('Answer score')):score_by_option[opt].append(float(x['Answer score']))
        for opt,cnt in option_counts.items():
            distrows.append([org,sec,qid,qtext,'Multi-select' if multi else 'Single-choice',opt,mean(score_by_option[opt]) if score_by_option[opt] else None,cnt,cnt/respondents if respondents else 0,respondents])
    qheaders=['Organisation name','Section','Question ID','Question','Respondents','Scored question','Most common answer','Most common count','Consensus','Average score','Median score','Minimum score','Maximum score','Score range','Standard deviation','Normalised score','Agreement','Review flag']
    rewrite(wb['Question Summary'],qheaders,qrows)
    dheaders=['Organisation name','Standard section','Question ID','Question','Question type','Answer option','Answer score','Respondents selecting','Percentage','Question respondents']
    rewrite(wb['Answer Distribution'],dheaders,distrows)

    # Section summary
    srows=[]; section_lookup={}
    for org in orgs:
        for sec in SCORED:
            subset=[r for r in qrows if r[0]==org and r[1]==sec and r[5]]
            if not subset: continue
            raw=[a for a in by_org_answers[org] if a['_section']==sec and num(a.get('Max score')) and a.get('Max score')>0 and num(a.get('Answer score'))]
            vals=[float(a['Answer score']) for a in raw]
            respondents=len({norm_text(a.get('Respondent ID')) for a in raw})
            cons=[r[8] for r in subset]
            mx=max([float(a['Max score']) for a in raw] or [4])
            row=[org,sec,respondents,len(subset),mean(vals),median(vals),min(vals),max(vals),mx,mean(vals)/mx,mean(cons),agreement(respondents,mean(cons))]
            srows.append(row); section_lookup[(org,sec)]=row
    sheaders=['Organisation name','Section','Respondents','Questions','Average score','Median score','Minimum score','Maximum score','Max possible','Normalised score','Average consensus','Agreement']
    rewrite(wb['Section Summary'],sheaders,srows)

    # Organisation Summary
    orows=[]
    for org in orgs:
        aa=by_org_answers[org]; ss=by_org_scores[org]
        ids={norm_text(a.get('Respondent ID')) for a in aa}
        meta=(ss[0] if ss else aa[0])
        overall_vals=[float(s['totalScore'])/100 for s in ss if num(s.get('totalScore'))]
        if not overall_vals:
            rv=[float(a['Answer score'])/float(a['Max score']) for a in aa if num(a.get('Answer score')) and num(a.get('Max score')) and a['Max score']>0]
            overall_vals=[mean(rv)] if rv else [0]
        overall=mean(overall_vals); avgscore=overall*4
        secvals={sec:section_lookup[(org,sec)][9] for sec in SCORED if (org,sec) in section_lookup}
        strongest=max(secvals,key=secvals.get) if secvals else ''; weakest=min(secvals,key=secvals.get) if secvals else ''
        qq=[r for r in qrows if r[0]==org]
        avgcons=mean([r[8] for r in qq]) if qq else 0
        reviews=sum(r[17]=='Review' for r in qq)
        n=len(ids); interp='Single respondent: perception only' if n<2 else ('Directional: small sample' if n<3 else 'Multi-respondent view')
        row=[org,norm_text(meta.get('Organisation type')),n,len({norm_text(a.get('Department')) for a in aa if norm_text(a.get('Department'))}),len({norm_text(a.get('Role level')) for a in aa if norm_text(a.get('Role level'))}),norm_text(meta.get('Organisation size')),norm_text(meta.get('PCDS sector')),norm_text(max([a.get('Submitted at') for a in aa],key=lambda x:norm_text(x),default='')),avgscore,overall,strongest,weakest,avgcons,reviews,agreement(n,avgcons),interp,maturity(overall),distance(overall)]
        orows.append(row)
    orows.sort(key=lambda r:(-r[2],-r[9],r[0].casefold()))
    oheaders=['Organisation name','Organisation type','Respondents','Departments represented','Role levels represented','Organisation size','Sector','Latest submission','Average score','Overall score','Strongest section','Weakest section','Average consensus','Questions for review','Agreement','Interpretation','Maturity tier','Distance to next tier']
    rewrite(wb['Organisation Summary'],oheaders,orows)

    # Lists and dropdowns
    lists=wb['Lists']
    for r in range(1,max(lists.max_row,len(orgs))+1):lists.cell(r,1).value=None
    for r,o in enumerate(orgs,1):lists.cell(r,1,o)
    lists.sheet_state='hidden'
    if 'OrganisationList' in wb.defined_names:del wb.defined_names['OrganisationList']
    wb.defined_names.add(DefinedName('OrganisationList',attr_text=f"'Lists'!$A$1:$A${len(orgs)}"))
    for sn,cell in [('Dashboard','B4'),('Organisation Report','B3')]:
        ws=wb[sn]; ws[cell]=ws[cell].value if ws[cell].value in orgs else orgs[0]
        ws.data_validations.dataValidation.clear()
        dv=DataValidation(type='list',formula1='=OrganisationList',allow_blank=False); dv.showDropDown=False
        ws.add_data_validation(dv);dv.add(ws[cell])
        ws[cell].fill=PatternFill('solid',fgColor='E2F0D9');ws[cell].font=Font(bold=True,color='008000')

    # Priority Detail hidden helper, used by Organisation Report top-five formulas
    pd=wb['Priority Detail'] if 'Priority Detail' in wb.sheetnames else wb.create_sheet('Priority Detail')
    pd.sheet_state='hidden'; pheaders=['Organisation','Question ID','Question','Most common answer','Normalised score','Agreement','Review flag','Priority rank','Lookup key']
    prows=[]
    byorg=defaultdict(list)
    for r in qrows:
        if r[5] and r[15] is not None:byorg[r[0]].append(r)
    for org in orgs:
        arr=sorted(byorg[org],key=lambda r:(0 if r[17]=='Review' else 1,r[15],r[2]))
        for rank,r in enumerate(arr,1):prows.append([org,r[2],r[3],r[6],r[15],r[16],r[17],rank,f'{org}|{rank}'])
    rewrite(pd,pheaders,prows,start=2)
    pd.sheet_state='hidden'

    # Fix Organisation Report priority formulas and section agreements for expandable data
    rep=wb['Organisation Report']
    for rr in range(11,18):
        rep.cell(rr,3,f'=IFERROR(LOOKUP(2,1/((\'Section Summary\'!$A:$A=$B$3)*(\'Section Summary\'!$B:$B=$A{rr})),\'Section Summary\'!$L:$L),"")')
    # Locate priority header row
    ph=next((r for r in range(1,rep.max_row+1) if rep.cell(r,1).value=='Question ID'),None)
    if ph:
        for rank,outrow in enumerate(range(ph+1,ph+6),1):
            helper=f'MATCH($B$3&"|{rank}",\'Priority Detail\'!$I:$I,0)'
            for c,col in enumerate('BCDEFG',1):rep.cell(outrow,c,f'=IFERROR(INDEX(\'Priority Detail\'!${col}:${col},{helper}),"")')

    # Dashboard and report formulas are preserved from template. Set selected org and force recalc.
    wb.calculation.fullCalcOnLoad=True;wb.calculation.forceFullCalc=True;wb.calculation.calcMode='auto'
    wb.active=wb.sheetnames.index('Dashboard')
    wb.save(output_file)
    # Reopen once and write standard data validations after the main workbook save.
    # This avoids legacy x14 validation extensions in template files suppressing new rules.
    final_wb=load_workbook(output_file,data_only=False)
    for sn,cell in [('Dashboard','B4'),('Organisation Report','B3')]:
        final_ws=final_wb[sn]
        final_ws.data_validations.dataValidation.clear()
        final_dv=DataValidation(type='list',formula1='=OrganisationList',allow_blank=False)
        final_dv.showDropDown=False
        final_ws.add_data_validation(final_dv); final_dv.add(final_ws[cell])
    final_wb.save(output_file)
    print(f'Created {output_file} from {len(answers)} answer rows, {len(scores)} score rows, {len(orgs)} organisations.')

# ═══════════════════════════════════════════════════════════════════════════
# CONFIG — edit these paths
# ═══════════════════════════════════════════════════════════════════════════

INPUT_FILE = "SARI_Results_2026-08-04-00-36-28.xlsx"
OUTPUT_FILE = "SARI_Organisation.xlsx"
TEMPLATE_FILE = str(Path(__file__).with_name("report_template.xlsx"))


def main():
    print(f"Reading: {INPUT_FILE}")
    build(INPUT_FILE, TEMPLATE_FILE, OUTPUT_FILE)
    print(f"Saved: {OUTPUT_FILE}")


if __name__ == "__main__":
    main()

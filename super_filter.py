#!/usr/bin/env python3
"""Generate the complete SARI_Organisation workbook from a raw SARI export.

Two stages, deliberately separated:

    compute(input_file)   -> every derived table, as plain Python rows. Fast (~1s).
    build(input_file,...) -> writes those rows into report_template.xlsx. Slow (~70s),
                             because openpyxl copies row styling cell by cell.

The GUI displays what compute() returns and exports what build() writes, so the table
on screen and the Organisation Summary sheet in the file can never disagree.

The supplied report_template.xlsx controls visual formatting, charts, page setup and
dashboard layout. This script only replaces data and repairs the organisation dropdowns.
"""
from __future__ import annotations
import argparse, math, shutil, sys
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
MULTI_SELECT={'background_2','background_3','background_4'}

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

def split_multi(answer):
    """Split a multi-select answer on commas, but NOT on commas inside brackets.

    'Shared infrastructure (e.g., cloud, computing power)' is ONE option, not three.
    A plain split() shredded it across 39 organisations and silently undercounted it.
    """
    text=norm_text(answer); parts=[]; buf=[]; depth=0
    for ch in text:
        if ch in '([{': depth+=1
        elif ch in ')]}': depth=max(0,depth-1)
        if ch==',' and depth==0:
            parts.append(''.join(buf)); buf=[]
        else: buf.append(ch)
    parts.append(''.join(buf))
    return [p.strip() for p in parts if p.strip()]

def copy_row_style(ws,src,dst,max_col):
    for c in range(1,max_col+1):
        ws.cell(dst,c)._style=copy(ws.cell(src,c)._style)
        ws.cell(dst,c).number_format=ws.cell(src,c).number_format
        ws.cell(dst,c).alignment=copy(ws.cell(src,c).alignment)
    ws.row_dimensions[dst].height=ws.row_dimensions[src].height

def rewrite(ws, headers, rows, start=5, add_filter=True):
    max_col=len(headers)
    for r in range(start, max(ws.max_row,start)+1):
        for c in range(1,max(ws.max_column,max_col)+1): ws.cell(r,c).value=None
    for c,h in enumerate(headers,1): ws.cell(start-1,c,h)
    for i,row in enumerate(rows,start):
        if i!=start: copy_row_style(ws,start,i,max_col)
        for c,v in enumerate(row,1): ws.cell(i,c,v)
    # Drop trailing rows left over from a longer previous run, so the sheet ends
    # where the data ends instead of trailing dozens of blank styled rows.
    last=start+len(rows)-1
    if ws.max_row>last: ws.delete_rows(last+1,ws.max_row-last)
    span=f'A{start-1}:{ws.cell(start-1,max_col).column_letter}{last}'

    # Four sheets carry an Excel Table (ListObject) from the template. Two things
    # about them cause the "Excel was able to open the file by repairing or removing
    # the unreadable content" dialog, and both are handled here:
    #
    #   1. A Table owns its own autoFilter. Adding a second, sheet-level autoFilter
    #      to the same sheet is invalid, so Excel repairs the file by deleting the
    #      Table. This fired on Section Summary on every single export.
    #   2. The template's table refs are frozen to the row counts of the dataset it
    #      was built from (A4:L872 and friends). Any survey with a different number
    #      of rows leaves the ref pointing past the data. openpyxl does not move
    #      them, and delete_rows above does not either, so re-point them explicitly.
    tables=list(getattr(ws,'tables',{}).values())
    if tables:
        for tbl in tables:
            tbl.ref=span
            if tbl.autoFilter is not None: tbl.autoFilter.ref=span
        ws.auto_filter.ref=None
    elif add_filter:
        ws.auto_filter.ref=span

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

# ═══════════════════════════════════════════════════════════════════════════
# STAGE 1 — the numbers. No workbook, no openpyxl writing, no template.
# ═══════════════════════════════════════════════════════════════════════════

RAW_HEADERS=['Respondent ID','Submitted at','Standard section','Question #','Question ID','Question','Answer','Answer value','Answer score','Max score','Participant name','Job title','Organisation name','Organisation type','Organisation size','Stakeholder category','PCDS sector','District','Role level','Department','Age band','Part of group','Parent company']
QUESTION_HEADERS=['Organisation name','Section','Question ID','Question','Respondents','Scored question','Most common answer','Most common count','Consensus','Average score','Median score','Minimum score','Maximum score','Score range','Standard deviation','Normalised score','Agreement','Review flag']
DIST_HEADERS=['Organisation name','Standard section','Question ID','Question','Question type','Answer option','Answer score','Respondents selecting','Percentage','Question respondents']
SECTION_HEADERS=['Organisation name','Section','Respondents','Questions','Average score','Median score','Minimum score','Maximum score','Max possible','Normalised score','Average consensus','Agreement']
ORG_HEADERS=['Organisation name','Organisation type','Respondents','Departments represented','Role levels represented','Organisation size','Sector','Latest submission','Average score','Overall score','Strongest section','Weakest section','Average consensus','Questions for review','Agreement','Interpretation','Maturity tier','Distance to next tier']
PRIORITY_HEADERS=['Organisation','Question ID','Question','Most common answer','Normalised score','Agreement','Review flag','Priority rank','Lookup key']

def compute(input_file):
    """Read a raw SARI export and derive every output table.

    Returns a dict. Each table is a plain list of rows whose order matches the
    matching *_HEADERS constant above, so a caller can display it or write it
    without knowing anything about Excel.
    """
    ah,answers,sh,scores=read_source(input_file)

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
    raw_rows=[]
    for a in answers:
        raw_rows.append([a.get(h) if h!='Standard section' else a.get('_section','') for h in RAW_HEADERS])
    raw_rows.sort(key=lambda r:(norm_text(r[12]).casefold(),norm_text(r[0]),norm_text(r[4])))

    # Question summary and answer distribution
    qgroups=defaultdict(list)
    for a in answers:
        org=norm_text(a.get('Organisation name')); qid=norm_text(a.get('Question ID'))
        if org and qid:qgroups[(org,a['_section'],qid,norm_text(a.get('Question')))].append(a)
    qrows=[]; distrows=[]
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
        qrows.append([org,sec,qid,qtext,respondents,scored,mode,mode_count,consensus,av,med,mn,ma,rng,sd,ns,agr,review])
        # Full answer distribution. Multi-select options are split bracket-aware.
        multi=qid in MULTI_SELECT
        option_counts=Counter(); score_by_option=defaultdict(list)
        for x in grp:
            opts=split_multi(x.get('Answer')) if multi else [norm_text(x.get('Answer'))]
            for opt in filter(None,opts):
                option_counts[opt]+=1
                if num(x.get('Answer score')):score_by_option[opt].append(float(x['Answer score']))
        for opt,cnt in option_counts.items():
            distrows.append([org,sec,qid,qtext,'Multi-select' if multi else 'Single-choice',opt,mean(score_by_option[opt]) if score_by_option[opt] else None,cnt,cnt/respondents if respondents else 0,respondents])

    # Section summary
    srows=[]; section_lookup={}
    for org in orgs:
        for sec in SCORED:
            subset=[r for r in qrows if r[0]==org and r[1]==sec and r[5]]
            if not subset: continue
            raw=[a for a in by_org_answers[org] if a['_section']==sec and num(a.get('Max score')) and a.get('Max score')>0 and num(a.get('Answer score'))]
            vals=[float(a['Answer score']) for a in raw]
            if not vals: continue
            respondents=len({norm_text(a.get('Respondent ID')) for a in raw})
            cons=[r[8] for r in subset]
            mx=max([float(a['Max score']) for a in raw] or [4])
            row=[org,sec,respondents,len(subset),mean(vals),median(vals),min(vals),max(vals),mx,mean(vals)/mx,mean(cons),agreement(respondents,mean(cons))]
            srows.append(row); section_lookup[(org,sec)]=row

    # Organisation Summary
    orows=[]
    for org in orgs:
        aa=by_org_answers[org]; ss=by_org_scores[org]
        if not aa and not ss: continue
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
        orows.append([org,norm_text(meta.get('Organisation type')),n,len({norm_text(a.get('Department')) for a in aa if norm_text(a.get('Department'))}),len({norm_text(a.get('Role level')) for a in aa if norm_text(a.get('Role level'))}),norm_text(meta.get('Organisation size')),norm_text(meta.get('PCDS sector')),norm_text(max([a.get('Submitted at') for a in aa],key=lambda x:norm_text(x),default='')),avgscore,overall,strongest,weakest,avgcons,reviews,agreement(n,avgcons),interp,maturity(overall),distance(overall)])
    orows.sort(key=lambda r:(-r[2],-r[9],r[0].casefold()))

    # Priority Detail hidden helper, used by Organisation Report top-five formulas
    prows=[]; byorg=defaultdict(list)
    for r in qrows:
        if r[5] and r[15] is not None:byorg[r[0]].append(r)
    for org in orgs:
        arr=sorted(byorg[org],key=lambda r:(0 if r[17]=='Review' else 1,r[15],r[2]))
        for rank,r in enumerate(arr,1):prows.append([org,r[2],r[3],r[6],r[15],r[16],r[17],rank,f'{org}|{rank}'])

    return {
        'orgs':orgs,
        'respondents':len({norm_text(a.get('Respondent ID')) for a in answers}),
        'answer_rows':len(answers),
        'score_rows':len(scores),
        'raw':raw_rows,
        'question':qrows,
        'distribution':distrows,
        'section':srows,
        'organisation':orows,
        'priority':prows,
    }

# ═══════════════════════════════════════════════════════════════════════════
# STAGE 2 — the workbook. Takes what compute() produced and writes the template.
# ═══════════════════════════════════════════════════════════════════════════

def build(input_file,template_file,output_file,data=None,progress=None):
    """Write a full SARI workbook. Pass `data` from a previous compute() to skip
    recalculating. `progress` is an optional callable(str) for a status line."""
    say=progress or (lambda m: None)
    if data is None:
        say('Reading and calculating...')
        data=compute(input_file)
    orgs=data['orgs']
    if not orgs:
        raise ValueError('No organisations found in the input file')

    say('Opening template...')
    shutil.copy2(template_file,output_file)
    wb=load_workbook(output_file,data_only=False)
    for bad in ['Organisation Type','Section Benchmark']:
        if bad in wb.sheetnames: del wb[bad]

    say('Writing Raw Answers...')
    rewrite(wb['Raw Answers'],RAW_HEADERS,data['raw'],add_filter=False)
    say('Writing Question Summary...')
    rewrite(wb['Question Summary'],QUESTION_HEADERS,data['question'],add_filter=False)
    say('Writing Answer Distribution...')
    rewrite(wb['Answer Distribution'],DIST_HEADERS,data['distribution'],add_filter=False)
    say('Writing Section Summary...')
    rewrite(wb['Section Summary'],SECTION_HEADERS,data['section'])
    say('Writing Organisation Summary...')
    rewrite(wb['Organisation Summary'],ORG_HEADERS,data['organisation'])

    say('Rebuilding dropdowns...')
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

    say('Writing priority helper...')
    pdet=wb['Priority Detail'] if 'Priority Detail' in wb.sheetnames else wb.create_sheet('Priority Detail')
    pdet.sheet_state='hidden'
    rewrite(pdet,PRIORITY_HEADERS,data['priority'],start=2)
    pdet.sheet_state='hidden'

    # Organisation Report: section agreement + top-five priority formulas.
    # INDEX on an empty helper cell returns 0, not "", so each is wrapped to stay blank.
    rep=wb['Organisation Report']
    for rr in range(11,18):
        rep.cell(rr,3,f'=IFERROR(LOOKUP(2,1/((\'Section Summary\'!$A:$A=$B$3)*(\'Section Summary\'!$B:$B=$A{rr})),\'Section Summary\'!$L:$L),"")')
    ph=next((r for r in range(1,rep.max_row+1) if rep.cell(r,1).value=='Question ID'),None)
    if ph:
        for rank,outrow in enumerate(range(ph+1,ph+6),1):
            helper=f'MATCH($B$3&"|{rank}",\'Priority Detail\'!$I:$I,0)'
            for c,col in enumerate('BCDEFG',1):
                idx=f'INDEX(\'Priority Detail\'!${col}:${col},{helper})'
                rep.cell(outrow,c,f'=IFERROR(IF({idx}="","",{idx}),"")')

    say('Saving workbook...')
    wb.calculation.fullCalcOnLoad=True;wb.calculation.forceFullCalc=True;wb.calculation.calcMode='auto'
    wb.active=wb.sheetnames.index('Dashboard')
    wb.save(output_file)

    # Reopen once and write standard data validations after the main workbook save.
    # This avoids legacy x14 validation extensions in template files suppressing new rules.
    say('Finalising dropdowns...')
    final_wb=load_workbook(output_file,data_only=False)
    for sn,cell in [('Dashboard','B4'),('Organisation Report','B3')]:
        final_ws=final_wb[sn]
        final_ws.data_validations.dataValidation.clear()
        final_dv=DataValidation(type='list',formula1='=OrganisationList',allow_blank=False)
        final_dv.showDropDown=False
        final_ws.add_data_validation(final_dv); final_dv.add(final_ws[cell])
    final_wb.save(output_file)
    say('Done')
    return data

# ═══════════════════════════════════════════════════════════════════════════
# PATHS — resolved so they also work inside a PyInstaller bundle
# ═══════════════════════════════════════════════════════════════════════════

def resource_path(name):
    """Locate a data file whether running from source or from a frozen bundle.

    A frozen app unpacks its data flat into sys._MEIPASS, NOT next to the module,
    so Path(__file__).with_name() silently points at a file that is not there.
    From source the same files live beside this module or under assets/.
    """
    here=Path(__file__).resolve().parent
    roots=[]
    base=getattr(sys,'_MEIPASS',None)
    if base: roots.append(Path(base))
    roots += [here, here/'assets']
    for r in roots:
        p=r/name
        if p.exists(): return str(p)
    return str(here/name)

INPUT_FILE = "SARI_Results_2026-08-04-00-36-28.xlsx"
OUTPUT_FILE = "SARI_Organisation.xlsx"
TEMPLATE_FILE = resource_path("report_template.xlsx")


def main():
    ap=argparse.ArgumentParser(description='Build the SARI organisation workbook.')
    ap.add_argument('input',nargs='?',default=INPUT_FILE,help='raw SARI export .xlsx')
    ap.add_argument('-o','--output',default=OUTPUT_FILE,help='workbook to write')
    ap.add_argument('-t','--template',default=TEMPLATE_FILE,help='report template .xlsx')
    args=ap.parse_args()
    print(f'Reading: {args.input}')
    build(args.input,args.template,args.output,progress=lambda m:print(f'  {m}'))
    print(f'Saved: {args.output}')


if __name__ == "__main__":
    main()

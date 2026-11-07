"""Build the 8-model Excel report with openpyxl (a public PyPI dependency).

Run from the project directory after run_nigeria_nowcast.py produces Results/*.csv.
No artifact_tool, internet connection, or Microsoft Excel installation needed.
"""
from __future__ import annotations
import csv
import json
import math
from pathlib import Path
from openpyxl import Workbook
from openpyxl.workbook.properties import CalcProperties
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

ROOT = Path(__file__).resolve().parent
D = ROOT / 'Results'
NAVY='17314B'; WHITE='FFFFFF'; PALE='F3F7FB'; BLUE='E7F1FB'; INK='24344A'

def records(filename):
    with (D / filename).open(encoding='utf-8-sig', newline='') as f:
        return list(csv.DictReader(f))

def numeric(value):
    try:
        x=float(value)
        return x if math.isfinite(x) else None
    except (TypeError,ValueError):
        return None

def add_table(wb, name, headers, data, widths=None):
    ws=wb.create_sheet(name)
    ws.append(headers)
    for item in data: ws.append(item)
    ws.freeze_panes='A2'
    ws.auto_filter.ref=ws.dimensions
    for c in ws[1]:
        c.fill=PatternFill('solid', fgColor=NAVY)
        c.font=Font(color=WHITE,bold=True)
        c.alignment=Alignment(wrap_text=True,vertical='center')
    ws.row_dimensions[1].height=31
    for col in range(1,len(headers)+1):
        ws.column_dimensions[get_column_letter(col)].width=(widths or {}).get(col, 20)
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.alignment=Alignment(vertical='top',wrap_text=False)
            if cell.row % 2 == 0:
                cell.fill=PatternFill('solid',fgColor=PALE)
    return ws

def run():
    with (D/'forecast.json').open(encoding='utf-8') as f: summary=json.load(f)
    models=records('model_forecasts.csv')
    scores=records('backtest_scores.csv')
    history=records('historical_forecasts.csv')
    predictors=records('used_predictors.csv')
    coefficients=records('ols_coefficients.csv')
    wb=Workbook()
    wb.calculation=CalcProperties(calcMode='auto',fullCalcOnLoad=True,forceFullCalc=True)
    main=wb.active; main.title='Executive Summary'
    main.merge_cells('A1:H1'); main['A1']='NIGERIA REAL GDP | 8-MODEL NOWCAST'
    main['A1'].fill=PatternFill('solid',fgColor=NAVY)
    main['A1'].font=Font(size=16,bold=True,color=WHITE)
    main.row_dimensions[1].height=36
    main['A3']='As of'; main['B3']=summary['as_of']
    main['C3']='Target quarter'; main['D3']=summary['target_quarter']
    main['E3']='Vintage'; main['F3']=summary['vintage']
    main['G3']='Eligible models'; main['H3']=summary['models_eligible']
    metrics=[
        (5,'Statistical ensemble (%)',summary['statistical_ensemble_pct']),
        (6,'Expert judgment adjustment (pp)',summary['expert_adjustment_pp']),
        (7,'Published candidate (%)', '=B5+B6'),
        (8,'Original PI50 lower (%)',summary['pi50_lower_pct']),
        (9,'Original PI50 upper (%)',summary['pi50_upper_pct']),
        (10,'CAPE adjusted lower (%)','=ROUND(ROUND(B8,2)*1.10,2)'),
        (11,'CAPE adjusted upper (%)','=ROUND(ROUND(B9,2)*0.90,2)'),
    ]
    for row,label,value in metrics:
        main.cell(row,1,label); main.cell(row,2,value)
        main.cell(row,2).number_format='0.0000' if row<10 else '0.00'
        for col in (1,2):
            main.cell(row,col).fill=PatternFill('solid',fgColor=BLUE if row in (7,10,11) else PALE)
            if row in (7,10,11): main.cell(row,col).font=Font(bold=True,color=NAVY)
    main['D5']='Best model, latest 8'; main['E5']=summary['best_single_model_last8']
    main['D6']='Best-model forecast (%)'; main['E6']=summary['best_single_model_forecast_pct']
    main['D7']='Best-model RMSE (pp)'; main['E7']=summary['best_model_rmse_pp']
    main['D8']='PI50 calibration N'; main['E8']=summary['pi50_calibration_n']
    main.merge_cells('A12:H12'); main['A12']='Individual forecasts and ensemble contributions';main['A12'].font=Font(bold=True,size=13,color=NAVY)
    headers=['Rank','Model','Forecast %','Weight','Contribution pp','RMSE 8Q (pp)','MAE 8Q (pp)','Status']
    for col,head in enumerate(headers,1):
        c=main.cell(13,col,head); c.fill=PatternFill('solid',fgColor=NAVY); c.font=Font(color=WHITE,bold=True)
    sorted_models=sorted(models,key=lambda z: int(z['rank_m3_recent8']) if z['rank_m3_recent8'] else 999)
    for idx,item in enumerate(sorted_models,14):
        row=[int(item['rank_m3_recent8']) if item['rank_m3_recent8'] else None,
             item['model'],numeric(item['prediction_pct']),numeric(item['weight']),
             f'=IF(OR(C{idx}="",D{idx}=""),0,C{idx}*D{idx})',
             numeric(item['rmse_last8_pp']),numeric(item['mae_last8_pp']),item['status']]
        for col,v in enumerate(row,1): main.cell(idx,col,v)
        for col in (3,5,6,7): main.cell(idx,col).number_format='0.0000'
        main.cell(idx,4).number_format='0.00%'
        if idx%2==0:
            for col in range(1,9): main.cell(idx,col).fill=PatternFill('solid',fgColor=PALE)
    last=13+len(sorted_models)
    total=last+2
    main.cell(total,2,'Weight totals')
    main.cell(total,4,f'=SUM(D14:D{last})').number_format='0.00%'
    main.cell(total,5,f'=SUM(E14:E{last})').number_format='0.0000'
    main.merge_cells(start_row=total+2,start_column=1,end_row=total+2,end_column=8)
    main.cell(total+2,1,'PI50 is retrospective; the narrower CAPE range is editorial, not a calibrated statistical interval.')
    main.cell(total+2,1).alignment=Alignment(wrap_text=True)
    main.row_dimensions[total+2].height=32
    for idx,width in enumerate((13,35,18,14,20,22,22,24),1): main.column_dimensions[get_column_letter(idx)].width=width
    main.freeze_panes='C14'
    chart=BarChart()
    chart.title='Recent individual-model RMSE (8Q)'
    chart.y_axis.title='RMSE (percentage points)'
    chart.add_data(Reference(main,min_col=6,min_row=13,max_row=last),titles_from_data=True)
    chart.set_categories(Reference(main,min_col=2,min_row=14,max_row=last))
    chart.width=24;chart.height=10
    main.add_chart(chart,'J3')

    best=[]
    for vintage in ('M1+14','M2+14','M3+14'):
        for nquarter in (8,16,24):
            choices=[x for x in scores if x['vintage']==vintage and int(x['window_quarters'])==nquarter and x['model']!='ENSEMBLE']
            choices.sort(key=lambda z:(float(z['rmse_pp']),float(z['mae_pp'])))
            for rank,x in enumerate(choices[:3],1):
                best.append([vintage,nquarter,rank,x['model'],numeric(x['rmse_pp']),numeric(x['mae_pp']),int(x['n'])])
    add_table(wb,'Best Models',['Information set','Window Q','Rank','Model','RMSE pp','MAE pp','N'],best,{1:18,2:16,3:12,4:34,5:18,6:18,7:10})
    model_rows=[[z['model'],z['predictors'],numeric(z['prediction_pct']),numeric(z['weight']),numeric(z['contribution_pp']),
                 int(z['rank_m3_recent8']) if z['rank_m3_recent8'] else None,
                 numeric(z['rmse_last8_pp']),numeric(z['mae_last8_pp']),z['status']] for z in models]
    s=add_table(wb,'All Model Forecasts',['Model','Predictors','Forecast %','Weight','Contribution pp','Rank 8Q','RMSE 8Q','MAE 8Q','Status'],model_rows,{1:34,2:44,3:18,4:16,5:22,6:13,7:17,8:17,9:20})
    for row in range(2,s.max_row+1):s.cell(row,4).number_format='0.00%'
    add_table(wb,'Backtest Scores',['Vintage','Model','Window Q','N','RMSE pp','MAE pp','Bias pp'],
              [[x['vintage'],x['model'],int(x['window_quarters']),int(x['n']),numeric(x['rmse_pp']),numeric(x['mae_pp']),numeric(x['bias_pp'])] for x in scores],
              {1:18,2:35,3:17,4:12,5:18,6:18,7:18})
    s=add_table(wb,'OLS Coefficients',['Term','Coefficient','Definition'],
              [[z['term'],numeric(z['coefficient']),'Intercept (estimated)' if z['term']=='Intercept' else 'Original source units'] for z in coefficients],
              {1:34,2:26,3:48})
    for row in range(2,s.max_row+1):s.cell(row,2).number_format='0.0000000000'
    descriptions={'gdp_latest_published':'Most recently published GDP growth (%)',
                  'cps_level_qtd_mean':'Nominal CPS level; original workbook units',
                  'pmi_qtd_mean':'Mean of available CBN Composite PMI observations',
                  'risk_qtd_mean':'Mean of available inflation risk observations'}
    add_table(wb,'Model Inputs',['Feature','Available value','Description'],
              [[z['variable'],numeric(z['value']),descriptions.get(z['variable'],'')] for z in predictors],{1:38,2:24,3:78})
    add_table(wb,'Historical Errors',['Vintage','Forecast date','Quarter','GDP release date','Model','Actual %','Forecast %','Error pp'],
              [[x['vintage'],x['as_of_date'],x['quarter'],x['gdp_release_date'],x['model'],
                numeric(x['actual_pct']),numeric(x['prediction_pct']),numeric(x['error_pp'])] for x in history],
              {1:16,2:19,3:14,4:20,5:34,6:18,7:18,8:18})
    notes=[
        ('Retained models','Persistence; AR Ridge; Ridge Bridge; MIDAS-style Ridge; Dynamic Factor Proxy; Ridge; Elastic Net; OLS with intercept.'),
        ('Model inputs','Latest released GDP growth, CPS nominal level, CBN Composite PMI, inflation risk; some benchmarks use only GDP.'),
        ('No tree models','Random Forest and Histogram Boosting excluded from this 8-model package.'),
        ('OLS','OLS is fit with an intercept and raw CPS LEVEL, not CPS growth or momentum.'),
        ('Information availability','60-CALENDAR-DAY SCENARIO: every GDP quarter assumed available 60 days after its end; actual GDP publication dates intentionally ignored. Monthly release rules unchanged.'),
        ('Historical-vintage caveat','Historical monthly predictor observations are later-vintage proxies, not complete archival snapshots.'),
        ('Best model','Ranked by RMSE in matched trailing historical error windows; MAE used as a tie-break.'),
        ('Ensemble','Inverse squared RMSE over the last eight published-outcome errors within the same information set.'),
        ('PI50','Retrospectively calibrated empirical residual interval; prospective coverage not established.'),
        ('CAPE reporting range','Displayed 2-decimal PI50 bounds adjusted lower x1.10 and upper x0.90; editorial not calibrated PI50.'),
        ('Expert judgment','Separate user-entered point adjustment plus reason; currently 0 pp.'),
        ('Financial data','Original master workbook not overwritten by model runner.'),
    ]
    s=add_table(wb,'Methodology and Caveats',['Audit topic','Method and qualification'],notes,{1:31,2:115})
    for row in range(2,s.max_row+1):
        s.cell(row,2).alignment=Alignment(wrap_text=True,vertical='center')
        s.row_dimensions[row].height=37
    add_table(wb,'PI50 Audit',['Field','Value','Clarification'],[
        ['Published candidate',summary['published_candidate_forecast_pct'],'Statistical ensemble plus any recorded expert adjustment'],
        ['Raw ensemble',summary['statistical_ensemble_pct'],'Model-only forecast'],
        ['Expert adjustment',summary['expert_adjustment_pp'],summary.get('expert_reason') or 'None'],
        ['Calibration N',summary['pi50_calibration_n'],'Available historical forecast errors'],
        ['Error radius',summary['pi50_radius_pp'],'Retrospective empirical-error radius (pp)'],
        ['PI50 lower',summary['pi50_lower_pct'],'Original statistical lower endpoint'],
        ['PI50 upper',summary['pi50_upper_pct'],'Original statistical upper endpoint'],
        ['CAPE lower',summary['reporting_range_lower_pct'],'Displayed lower x1.10, editor adjusted'],
        ['CAPE upper',summary['reporting_range_upper_pct'],'Displayed upper x0.90, editor adjusted'],
        ['Interval validation','NOT PROSPECTIVELY VALIDATED','Actual historical vintage coverage missing'],
    ],{1:34,2:30,3:100})
    D.mkdir(exist_ok=True)
    outfile=D/'Nigeria_GDP_8_Model_Results.xlsx'
    wb.save(outfile)
    print('Excel workbook created:', outfile)
    print('Sheets:', ', '.join(wb.sheetnames))
    return outfile

if __name__=='__main__':run()

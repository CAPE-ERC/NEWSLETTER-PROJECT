"""Eight-model Nigeria GDP nowcast: seven original linear/factor families + OLS (with intercept).
All multivariable estimators use latest published GDP growth, nominal CPS level,
CBN Composite PMI, inflation risk. MIDAS-style model additionally uses released
month-specific observations of those same three monthly series. Historical
errors use release-gated outcomes, but historical predictor values are final
vintage proxies, not actual archived source vintages.
"""
from __future__ import annotations
import argparse,csv,hashlib,json,math,os,sys,warnings
from datetime import date,timedelta,datetime,timezone
from decimal import Decimal, ROUND_HALF_UP
from pathlib import Path
import numpy as np
import pandas as pd
from sklearn.pipeline import Pipeline,make_pipeline
from sklearn.preprocessing import StandardScaler
from sklearn.impute import SimpleImputer
from sklearn.decomposition import PCA
from sklearn.linear_model import LinearRegression,Ridge,ElasticNet
ROOT=Path(__file__).resolve().parent
sys.path.insert(0,str(ROOT/'engine'))
MASTER=ROOT.parent.parent/'Main Data'/'MultipleFrequencyModelData.xlsx'
RESULTS=ROOT/'Results'
FOUR=['gdp_latest_published','cps_level_qtd_mean','pmi_qtd_mean','risk_qtd_mean']
GDP=['gdp_latest_published','gdp_age_quarters','gdp_lag1','gdp_lag2','gdp_lag3']
SERIES=['cps_level','pmi','risk']
NAMES=['Persistence','AR Ridge','Ridge Bridge','MIDAS-style Ridge','Dynamic Factor Proxy','Ridge','Elastic Net','OLS with intercept']
MIN_TRAIN=20
MIN_WEIGHT_HISTORY=8

def displayed_pct(value):
 """Round a percent-number (e.g. 3.789 -> 3.79), not a proportion."""
 return Decimal(str(value)).quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)

def adjusted_reporting_range(lower, upper):
 """Apply +/-10% to the displayed PI50 bounds, then round to 0.01 pp.
 This is an editorial judgment range, not a recalibrated PI50.
 """
 if not finite(lower) or not finite(upper):return None,None
 low=(displayed_pct(lower)*Decimal('1.10')).quantize(Decimal('0.01'),rounding=ROUND_HALF_UP)
 high=(displayed_pct(upper)*Decimal('0.90')).quantize(Decimal('0.01'),rounding=ROUND_HALF_UP)
 return float(low),float(high)

def finite(x):
 try:return bool(np.isfinite(float(x)))
 except(TypeError,ValueError):return False

def sha256(p):
 h=hashlib.sha256()
 with open(p,'rb') as f:
  for chunk in iter(lambda:f.read(1<<20),b''):h.update(chunk)
 return h.hexdigest()

def write_csv(p,rows,cols):
 p.parent.mkdir(parents=True,exist_ok=True)
 with open(p,'w',encoding='utf-8-sig',newline='') as f:
  w=csv.DictWriter(f,fieldnames=cols,extrasaction='ignore');w.writeheader();w.writerows(rows)

def metrics(a):
 e=np.asarray(a,float)
 return {'n':len(e),'rmse_pp':float(np.sqrt(np.mean(e*e))),'mae_pp':float(np.mean(np.abs(e))),'bias_pp':float(np.mean(e))}

def make_panel(a,quarters,months):
 import release_aware
 release_aware.MONTHLY_SERIES['cps_level']='CPS'
 p=a.build_monthly_predictors(months)
 for item in months:p[(item['year'],item['month'])]['cps_level']=a.num(item.get('CPS'))
 names=a.select_predictors(p)
 if 'cps_level' not in names:raise RuntimeError('CPS nominal level missing')
 gm={a.quarter_index(r['year'],r['quarter']):r['growth'] for r in quarters}
 rows=[]
 for q in quarters:
  for v in (1,2,3):
   origin=a.as_of_vintage(q['year'],q['quarter'],v)+timedelta(days=14)
   rows.append(a.build_feature_row(q['year'],q['quarter'],v,names,p,gm,growth=q['growth'],origin=origin))
 return pd.DataFrame(rows),p,names,gm

def features_for(name,t):
 if name=='Persistence':return []
 if name=='AR Ridge':return GDP
 if name=='MIDAS-style Ridge':
  # Only release-available month-specific signals enter; no future-month imputation.
  extra=[f'{s}_m{m}' for s in SERIES for m in (1,2,3) if finite(t.get(f'{s}_m{m}'))]
  lag=[f'{s}_prev_q_mean' for s in SERIES if finite(t.get(f'{s}_prev_q_mean'))]
  return FOUR+extra+lag
 return FOUR

def predict(name,train,t):
 if name=='Persistence':return (float(t['gdp_latest_published']),{}) if finite(t.get('gdp_latest_published')) else (math.nan,{})
 if name not in ('AR Ridge',) and any(not finite(t.get(f)) for f in FOUR):
  return math.nan,{'status':'not_yet_published_four_signals'}
 feats=features_for(name,t)
 if name=='AR Ridge':
  if not finite(t.get('gdp_latest_published')):return math.nan,{'status':'no_released_GDP'}
  model=Ridge(alpha=2)
 elif name=='OLS with intercept':model=LinearRegression(fit_intercept=True)
 elif name=='Ridge Bridge':model=Ridge(alpha=2)
 elif name=='MIDAS-style Ridge':model=Ridge(alpha=4)
 elif name=='Dynamic Factor Proxy':model=make_pipeline(PCA(n_components=min(3,len(feats))),Ridge(alpha=3))
 elif name=='Ridge':model=Ridge(alpha=5)
 elif name=='Elastic Net':model=ElasticNet(alpha=.05,l1_ratio=.2,max_iter=5000,random_state=42)
 else:raise KeyError(name)
 # Complete training cases on the four inputs for all four-indicator models;
 # MIDAS extras are median imputed from TRAINING ONLY, not from test future months.
 if name=='AR Ridge':
  relevant=['gdp_latest_published','growth']
  sub=train.dropna(subset=relevant).copy()
 else:sub=train.replace([np.inf,-np.inf],np.nan).dropna(subset=FOUR+['growth']).copy()
 if len(sub)<MIN_TRAIN:return math.nan,{'status':'insufficient_training','n':len(sub)}
 if name=='OLS with intercept' and (sub[FOUR].std(ddof=0).values<1e-10).any():return math.nan,{'status':'rank_deficiency_risk'}
 if name in ('MIDAS-style Ridge','AR Ridge'):
  # Remove historical features never published by the relevant origin.
  # The latest GDP and four required features cannot be dropped.
  required=FOUR if name=='MIDAS-style Ridge' else ['gdp_latest_published']
  feats=[f for f in feats if f in required or sub[f].replace([np.inf,-np.inf],np.nan).notna().sum()>=8]
 x=sub[feats].replace([np.inf,-np.inf],np.nan)
 if x.isna().all(axis=0).any():return math.nan,{'status':'entirely_missing_feature'}
 tx=pd.DataFrame([{k:t.get(k,math.nan) for k in feats}])[feats]
 if name!='AR Ridge' and tx[FOUR].isna().any(axis=None):return math.nan,{'status':'missing_required_inputs'}
 if name=='AR Ridge' and not finite(t.get('gdp_latest_published')):return math.nan,{}
 pre=[('impute',SimpleImputer(strategy='median',keep_empty_features=True))]
 pre.append(('scale',StandardScaler()))
 pre.append(('reg',model))
 pipe=Pipeline(pre)
 with warnings.catch_warnings():
  warnings.filterwarnings('ignore',category=UserWarning)
  pipe.fit(x,sub.growth.to_numpy(float))
  yhat=float(pipe.predict(tx)[0])
 meta={'n_complete':len(sub),'features':feats}
 if name=='OLS with intercept':
  reg=pipe.named_steps['reg'];scale=pipe.named_steps['scale'];coef=reg.coef_/scale.scale_
  adjusted_intercept=float(reg.intercept_-coef@scale.mean_)
  meta.update({'intercept':adjusted_intercept,'coefficients':{k:float(v) for k,v in zip(feats,coef)},'r2_in_sample':float(pipe.score(x,sub.growth))})
 return yhat,meta

def historical_weights(recs,origin,v,available,min_count=MIN_WEIGHT_HISTORY):
 """No same-quarter information leakage: scores require target GDP release date <= origin."""
 out={}
 for name in available:
  rec=[r for r in recs if r['vintage']==f'M{v}+14' and r['model']==name and date.fromisoformat(r['gdp_release_date'])<=origin]
  rec=sorted(rec,key=lambda x:x['as_of_date'])[-min_count:]
  if len(rec)>=min_count:
   rmse=math.sqrt(sum(r['error_pp']**2 for r in rec)/len(rec))
   if rmse>0:out[name]=1.0/(rmse**2)
 # No shrinking or cherry-picking OLS: every forecastable model has non-zero weight.
 # Before min score count across all models: equal-weight available models.
 if len(out)!=len(available):return {name:1/len(available) for name in available} if available else {}
 z=sum(out.values());return {name:out[name]/z for name in available}

def combine(pred,w):return sum(pred[k]*v for k,v in w.items()) if w else math.nan

def run(as_of,target_quarter,skip_excel=False,expert_adjustment_pp=0.,expert_reason=''):
 from release_aware import MONTHLY_SERIES
 if expert_adjustment_pp and not expert_reason.strip():raise ValueError('--expert-reason required with non-zero adjustment')
 if not MASTER.exists():raise FileNotFoundError(MASTER)
 os.environ['NIGERIA_MASTER_WORKBOOK']=str(MASTER)
 os.environ['NIGERIA_AS_OF_DATE']=as_of
 os.environ['NIGERIA_TARGET_QUARTER']=target_quarter
 import data_features as a
 orig=sha256(MASTER);now=date.fromisoformat(as_of)
 y,q=int(target_quarter[:4]),int(target_quarter[-1]);v=min(3,max(1,(now.year*12+now.month)-(y*12+(q-1)*3+1)+1))
 if a.GUARD.gdp_available(y,q,now):raise ValueError('Target GDP published already: this is not a nowcast.')
 months=a.parse_monthly_sheet(MASTER);quarters=a.parse_quarterly_sheet(MASTER)
 panel,p,names,gm=make_panel(a,quarters,months)
 test=pd.Series(a.build_feature_row(y,q,v,names,p,gm,origin=now))
 historical=[]
 for vintage in (1,2,3):
  group=panel[panel.vintage==vintage].sort_values('qidx')
  for _,t in group.tail(24).iterrows():
   origin=date.fromisoformat(t.origin_date)
   tr=group[(group.qidx<int(t.qidx)) & group.apply(lambda x:a.GUARD.gdp_available(int(x.year),int(x.quarter),origin),axis=1)]
   if len(tr)<MIN_TRAIN:continue
   predictions={}
   for name in NAMES:
    try:pred,_=predict(name,tr,t)
    except Exception as e:raise RuntimeError(f'{name} failed at {t.period} {origin}: {e}') from e
    if finite(pred):
     predictions[name]=pred
     historical.append({'vintage':f'M{vintage}+14','as_of_date':origin.isoformat(),'quarter':t.period,'gdp_release_date':a.GUARD.gdp_date(int(t.year),int(t.quarter)).isoformat(),'model':name,'actual_pct':float(t.growth),'prediction_pct':pred,'error_pp':pred-float(t.growth),'weight':None})
   # Ensemble weights only from prior published GDP outcomes; cannot see this target.
   weights=historical_weights(historical,origin,vintage,list(predictions))
   yp=combine(predictions,weights)
   if finite(yp):
    historical.append({'vintage':f'M{vintage}+14','as_of_date':origin.isoformat(),'quarter':t.period,'gdp_release_date':a.GUARD.gdp_date(int(t.year),int(t.quarter)).isoformat(),'model':'ENSEMBLE','actual_pct':float(t.growth),'prediction_pct':yp,'error_pp':yp-float(t.growth),'weight':None})
  # End period
 train=panel[(panel.vintage==v)&(panel.qidx<int(test.qidx))&panel.apply(lambda x:a.GUARD.gdp_available(int(x.year),int(x.quarter),now),axis=1)].sort_values('qidx')
 if len(train)<MIN_TRAIN:raise ValueError('Insufficient published historical GDP samples')
 preds={};meta={}
 for name in NAMES:
  try:pr,md=predict(name,train,test)
  except Exception as e:raise RuntimeError(f'Current {name}: {e}') from e
  preds[name]=pr;meta[name]=md
 available=[k for k in NAMES if finite(preds[k])]
 weights=historical_weights(historical,now,v,available)
 raw=combine(preds,weights)
 if not finite(raw):raise RuntimeError('No models available at origin')
 forecast=raw+expert_adjustment_pp
 # Finite-sample 50% absolute-error interval based on released historical
 # ensemble forecasts at the SAME vintage. Empirical/provisional, no true-vintage coverage.
 hist_ensemble=sorted([h for h in historical if h['vintage']==f'M{v}+14' and h['model']=='ENSEMBLE' and date.fromisoformat(h['gdp_release_date'])<=now],key=lambda h:h['as_of_date'])[-16:]
 errors=sorted(abs(h['error_pp']) for h in hist_ensemble)
 if len(errors)>=8:
  rank=min(len(errors),math.ceil((len(errors)+1)*.5));radius=float(errors[rank-1])
  lo,hi=forecast-radius,forecast+radius
 else:lo=hi=math.nan
 scores=[]
 for vintage in (1,2,3):
  for name in NAMES+['ENSEMBLE']:
   rr=sorted([x for x in historical if x['vintage']==f'M{vintage}+14' and x['model']==name],key=lambda x:x['as_of_date'])
   for count in (8,16,24):
    if len(rr)>=count:scores.append({'vintage':f'M{vintage}+14','model':name,'window_quarters':count,**metrics([x['error_pp'] for x in rr[-count:]])})
 curr=[s for s in scores if s['vintage']==f'M{v}+14' and s['window_quarters']==8 and s['model'] in available]
 ranking=sorted(curr,key=lambda s:(s['rmse_pp'],s['mae_pp']))
 best=ranking[0]['model'] if ranking else None
 rankdict={s['model']:i+1 for i,s in enumerate(ranking)}
 scoremap={s['model']:s for s in curr}
 output=[]
 for name in NAMES:
  out={'model':name,'predictors':'GDP only' if name=='Persistence' else 'GDP historical lags' if name=='AR Ridge' else 'GDP, CPS level, CBN PMI, inflation risk',
       'prediction_pct':float(preds[name]) if finite(preds[name]) else None,
       'weight':weights.get(name,0.),'contribution_pp':weights.get(name,0.)*preds[name] if name in weights else 0.,
       'rank_m3_recent8':rankdict.get(name),'rmse_last8_pp':scoremap[name]['rmse_pp'] if name in scoremap else None,
       'mae_last8_pp':scoremap[name]['mae_pp'] if name in scoremap else None,
       'status':'Active' if name in weights else 'Missing required released predictor'}
  output.append(out)
 write_csv(RESULTS/'model_forecasts.csv',output,['model','predictors','prediction_pct','weight','contribution_pp','rank_m3_recent8','rmse_last8_pp','mae_last8_pp','status'])
 write_csv(RESULTS/'backtest_scores.csv',scores,['vintage','model','window_quarters','n','rmse_pp','mae_pp','bias_pp'])
 write_csv(RESULTS/'historical_forecasts.csv',historical,['vintage','as_of_date','quarter','gdp_release_date','model','actual_pct','prediction_pct','error_pp','weight'])
 vals=[{'variable':x,'value':float(test.get(x)) if finite(test.get(x)) else None} for x in FOUR]
 write_csv(RESULTS/'used_predictors.csv',vals,['variable','value'])
 coefs=[{'term':'Intercept','coefficient':meta.get('OLS with intercept',{}).get('intercept')}]
 for key,value in meta.get('OLS with intercept',{}).get('coefficients',{}).items():coefs.append({'term':key,'coefficient':value})
 write_csv(RESULTS/'ols_coefficients.csv',coefs,['term','coefficient'])
 range_low,range_high=adjusted_reporting_range(lo,hi)
 if range_low is not None and not (range_low<=forecast<=range_high):
  raise ValueError('Adjusted reporting range does not contain the point forecast; review the requested 10% reduction before publishing.')
 result={'as_of':as_of,'target_quarter':target_quarter,'vintage':f'M{v}',
  'model_predictions_pct':{k:(float(x) if finite(x) else None) for k,x in preds.items()},
  'ensemble_weights':weights,'statistical_ensemble_pct':raw,'expert_adjustment_pp':expert_adjustment_pp,'expert_reason':expert_reason,
  'published_candidate_forecast_pct':forecast,'reporting_range_lower_pct':range_low,'reporting_range_upper_pct':range_high,
  'reporting_range_definition':'Judgment-adjusted CAPE range: round both original PI50 bounds to 0.01 percentage point, increase displayed lower by 10%, reduce displayed upper by 10%, round to 0.01 pp; not a statistical PI50',
  'pi50_lower_pct':lo if finite(lo) else None,'pi50_upper_pct':hi if finite(hi) else None,
  'reporting_bound_adjustment_fraction':0.10,'reporting_input_rounding_decimals':2,
  'pi50_radius_pp':radius if len(errors)>=8 else None,'pi50_calibration_n':len(errors),
  'best_single_model_last8':best,'best_single_model_forecast_pct':preds[best] if best else None,
  'best_model_rmse_pp':ranking[0]['rmse_pp'] if ranking else None,'models_eligible':len(available),
  'cps_definition':'Raw workbook CPS level; not growth or momentum','ols_coefficients':meta.get('OLS with intercept',{}),
  'historical_method':'FINAL-VINTAGE PSEUDO-REAL-TIME, all quarterly GDP assumed published 60 calendar days after quarter-end; M1/M2/M3 origins at month-end+14 days; recorded GDP dates intentionally ignored',
  'gdp_release_assumption_days':60,'gdp_release_calendar_policy':'assumed_60_calendar_days_all_quarters_no_overrides',
  'weighting_method':'Inverse-squared RMSE of last 8 eligible historical errors by same vintage, including OLS, only GDP outcomes published by forecast origin; equal weights until 8 eligible errors per component',
  'interval_method':'Symmetric finite-sample median-absolute-error proxy from last 16 sequential ensemble errors already published; 50% nominal only, not independent calibration',
  'expert_judgment_policy':'No manual adjustment unless explicitly supplied with a written reason; raw ensemble reported separately.',
  'model_version':'eight_model_linear_factor_ols_2026_10_14'}
 RESULTS.mkdir(exist_ok=True)
 (RESULTS/'forecast.json').write_text(json.dumps(result,indent=2,allow_nan=False),encoding='utf-8')
 if sha256(MASTER)!=orig:raise RuntimeError('Master workbook changed')
 registry=ROOT/'prospective_forecast_registry.csv'
 cols=['as_of','target_quarter','vintage','model','forecast_pct','master_sha256','recorded_utc','status']
 prev=list(csv.DictReader(registry.open(encoding='utf-8-sig',newline=''))) if registry.exists() else []
 keys={(r['as_of'],r['target_quarter'],r['vintage'],r['model']) for r in prev}
 updates=[]
 for name,val in list(preds.items())+[('ENSEMBLE',forecast)]:
  if not finite(val) or (as_of,target_quarter,f'M{v}',name) in keys:continue
  updates.append(dict(as_of=as_of,target_quarter=target_quarter,vintage=f'M{v}',model=name,forecast_pct=val,master_sha256=orig,recorded_utc=datetime.now(timezone.utc).isoformat(),status='registered_before_GDP_release'))
 if updates:
  with registry.open('a',encoding='utf-8' if registry.exists() else 'utf-8-sig',newline='') as f:
   writer=csv.DictWriter(f,fieldnames=cols)
   if not prev:writer.writeheader()
   writer.writerows(updates)
 if not skip_excel:
  # openpyxl is installed by requirements_model.txt and works on ordinary Windows Python.
  # Report failure is surfaced instead of silently skipping the requested workbook.
  from subprocess import run as run_process
  run_process([sys.executable,str(ROOT/'build_report.py')],cwd=ROOT,check=True)
 print('FORECAST',round(forecast,6),'PI50',round(lo,6) if finite(lo) else None,round(hi,6) if finite(hi) else None)
 print('CAPE REPORTING RANGE',range_low,range_high,'(lower PI50 displayed x 1.10; upper PI50 displayed x 0.90; NOT statistical PI50)')
 print('BEST',best,'OLS WEIGHT',weights.get('OLS with intercept',0))
 print('MASTER SHA256',orig)
 return result

def main():
 ap=argparse.ArgumentParser()
 ap.add_argument('--as-of',default=date.today().isoformat())
 ap.add_argument('--target-quarter',help='YYYYQn; defaults to the earliest started quarter whose GDP is not yet released')
 ap.add_argument('--skip-excel',action='store_true')
 ap.add_argument('--expert-adjustment-pp',type=float,default=0.)
 ap.add_argument('--expert-reason',default='')
 args=ap.parse_args()
 target=args.target_quarter.upper() if args.target_quarter else infer_target_quarter(args.as_of)
 print('TARGET QUARTER',target,'(given)' if args.target_quarter else '(selected from the information date and the 60-day GDP release rule)')
 run(args.as_of,target,args.skip_excel,args.expert_adjustment_pp,args.expert_reason)

def infer_target_quarter(as_of):
 from release_aware import ReleaseGuard
 now=date.fromisoformat(as_of)
 current_q=(now.month-1)//3+1
 previous=(now.year-1,4) if current_q==1 else (now.year,current_q-1)
 guard=ReleaseGuard(ROOT/'engine'/'release_calendar.csv')
 year,quarter=previous if not guard.gdp_available(*previous,now) else (now.year,current_q)
 return f'{year}Q{quarter}'
if __name__=='__main__':main()

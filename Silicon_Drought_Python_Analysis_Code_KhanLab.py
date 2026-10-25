#!/usr/bin/env python3
"""20 species x 2 silicon x 2 drought, 10 variables. Source: Batch2-Silicate experiment values.xlsx.
Reproducible: python Silicon_Drought_Analysis.py /path/to/workbook.xlsx /path/to/results
Requires artifact_tool, numpy, scipy, scikit-learn, statsmodels, matplotlib, xgboost, python-pptx.
Scientific safeguards: replicate arrays independently by sheet; no fabricated plant-level joins;
20 species-treatment means for multivariate analyses; species-grouped 5-fold CV;
invalid negative Si flagged; undefined photosynthesis never reconstructed.
"""
import os,sys,csv,json,warnings,math,textwrap
from pathlib import Path
import numpy as np
from scipy import stats
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.colors import TwoSlopeNorm
from matplotlib.ticker import MaxNLocator
from sklearn.base import clone
from sklearn.pipeline import Pipeline
from sklearn.impute import SimpleImputer
from sklearn.preprocessing import StandardScaler
from sklearn.decomposition import PCA
from sklearn.model_selection import GroupKFold,GroupShuffleSplit,KFold,cross_validate
from sklearn.metrics import r2_score,mean_squared_error,mean_absolute_error
from sklearn.inspection import permutation_importance
from sklearn.linear_model import Ridge,Lasso,ElasticNet
from sklearn.ensemble import RandomForestRegressor,GradientBoostingRegressor
from sklearn.svm import SVR
from sklearn.cross_decomposition import CCA
from xgboost import XGBRegressor
from statsmodels.formula.api import ols
from statsmodels.stats.anova import anova_lm
from scipy.stats import tukey_hsd
from artifact_tool import Blob,SpreadsheetFile
from pptx import Presentation
from pptx.util import Inches,Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

INPUT=Path(sys.argv[1] if len(sys.argv)>1 else '/mnt/data/Batch2-Silicate experiment values.xlsx')
OUT=Path(sys.argv[2] if len(sys.argv)>2 else '/mnt/data/Silicon_Drought_Results')
FIG=OUT/'figures'; FIG.mkdir(parents=True,exist_ok=True)
SEED=42; rng=np.random.default_rng(SEED)
TREAT=['Control','Silicon','Drought','Silicon + Drought']
SH_NAMES=['Shoot length','Root length','Fressh Weight Shoot','Fressh Weight root','Dry weight shoot','Dry weight root','Stomatal conductance','VPD Leaf','Si Quantification Root','Si Quantification Soil']
NAMES=['Shoot length','Root length','Fresh shoot weight','Fresh root weight','Dry shoot weight','Dry root weight','Stomatal conductance','Leaf VPD','Root silicon','Soil silicon']
KEYS=['Shoot_Length','Root_Length','Fresh_Shoot','Fresh_Root','Dry_Shoot','Dry_Root','Stomatal_Conductance','VPD_Leaf','Root_Silicon','Soil_Silicon']
NR=[10,10,5,5,5,5,3,3,3,3]
COLORS=['#45546A','#168A83','#C65C57','#4C7EA8']
TEXT='#233146'; ACCENT='#147D7B'
plt.rcParams.update({'font.family':'DejaVu Sans','font.size':10,'axes.titlesize':13,'axes.labelsize':10,'axes.titleweight':'bold','axes.spines.top':False,'axes.spines.right':False,'figure.facecolor':'white','axes.facecolor':'white','savefig.facecolor':'white','axes.grid':False,'svg.fonttype':'none'})

def save_fig(name,fig,caption,notes='',w=10,h=5.5):
    fig.set_size_inches(w,h)
    fig.savefig(FIG/(name+'.png'),dpi=600,bbox_inches='tight',pad_inches=.17,facecolor='white')
    plt.close(fig)
    SLIDES.append((name,caption,notes))

def write_csv(name,header,rows):
    with open(OUT/name,'w',newline='',encoding='utf-8-sig') as f:
        wr=csv.writer(f);wr.writerow(header)
        for row in rows:
            wr.writerow([('' if isinstance(x,(float,np.floating)) and not np.isfinite(x) else x) for x in row])

def numeric(v):
    if isinstance(v,(float,int,np.floating,np.integer)) and not isinstance(v,bool):return float(v)
    try:return float(str(v).strip())
    except:return np.nan

print('Importing workbook using artifact_tool...',flush=True)
w=SpreadsheetFile.import_xlsx(Blob.load(str(INPUT)))
raw={sh:w.worksheets.get_item(sh).get_range('A1:AQ48').values for sh in SH_NAMES}
SPECIES=[str(r[0]).strip() for r in raw[SH_NAMES[0]][1:21]]
assert len(SPECIES)==20 and len(set(SPECIES))==20
# rows 1..20 (0-based), treatment blocks of width 10/5/3 in columns B onwards.
D=np.full((20,4,10,10),np.nan,dtype=float)
raw_obs=[]; qc=[]
for k in range(8):
    mat=raw[SH_NAMES[k]]
    assert [str(r[0]).strip() for r in mat[1:21]]==SPECIES,(SH_NAMES[k],'species order mismatch')
    n=NR[k]
    for s,sp in enumerate(SPECIES):
        for t,tr in enumerate(TREAT):
            for rep in range(n):
                value=numeric(mat[s+1][1+t*n+rep]); D[s,t,k,rep]=value
                raw_obs.append((sp,tr,NAMES[k],rep+1,value,'OK' if np.isfinite(value) else 'MISSING',SH_NAMES[k]))
                if not np.isfinite(value):qc.append((sp,tr,NAMES[k],rep+1,value,'Missing raw value'))
# Si sheets: three side-by-side blocks, species-label cells at 0,6,12, 6 rows per species.
si_species_seen={}
for k in (8,9):
    mat=raw[SH_NAMES[k]];seen=set()
    for strow in range(2,44,6):
        for col in (0,6,12):
            if strow>=len(mat):continue
            item=mat[strow][col]
            if not isinstance(item,str) or item.strip() not in SPECIES:continue
            sp=item.strip();seen.add(sp);s=SPECIES.index(sp)
            for t,tr in enumerate(TREAT):
                for rep in range(3):
                    v=numeric(mat[strow+1+rep][col+1+t])
                    validity='OK' if np.isfinite(v) and v>=0 else ('INVALID_NEGATIVE' if np.isfinite(v) else 'MISSING')
                    raw_obs.append((sp,tr,NAMES[k],rep+1,v,validity,SH_NAMES[k]))
                    if np.isfinite(v) and v<0:
                        qc.append((sp,tr,NAMES[k],rep+1,v,'Negative silicon content; excluded from analysis'))
                        v=np.nan
                    elif not np.isfinite(v):qc.append((sp,tr,NAMES[k],rep+1,v,'Missing raw value'))
                    D[s,t,k,rep]=v
    missing=sorted(set(SPECIES)-seen)
    si_species_seen[NAMES[k]]=sorted(seen)
    for sp in missing:qc.append((sp,'ALL',NAMES[k],None,np.nan,'Entire silicon trait missing for species'))

write_csv('00_raw_long_audit.csv',['Species','Treatment','Parameter','Replicate','Raw_value','Quality_status','Original_sheet'],raw_obs)
write_csv('00_data_quality_flags.csv',['Species','Treatment','Parameter','Replicate','Raw_value','Flag'],qc)
MEANS=np.full((20,4,10),np.nan)
SDS=np.full_like(MEANS,np.nan)
CNT=np.zeros((20,4,10),int)
for s in range(20):
 for t in range(4):
  for k in range(10):
   v=D[s,t,k,np.isfinite(D[s,t,k])]; CNT[s,t,k]=len(v)
   if len(v):MEANS[s,t,k]=np.mean(v)
   if len(v)>1:SDS[s,t,k]=np.std(v,ddof=1)
# derived total dry biomass; unlabelled mass units as recorded; add mean of two trait means (NOT per-plant paired).
TOT=MEANS[:,:,4]+MEANS[:,:,5]
WRITE_HEADER=['Species','Treatment']+[x+'_mean' for x in KEYS]+['Total_Dry_Weight_Mean','Root_Shoot_Dry_Ratio']+[x+'_n' for x in KEYS]
meanrows=[]
for s,sp in enumerate(SPECIES):
 for t,tr in enumerate(TREAT):
  meanrows.append([sp,tr,*MEANS[s,t,:],TOT[s,t],MEANS[s,t,5]/MEANS[s,t,4] if MEANS[s,t,4]>0 else np.nan,*CNT[s,t,:]])
write_csv('01_species_treatment_means.csv',WRITE_HEADER,meanrows)
print('Data assembled: 20 species, 4 treatments, 10 traits. Missing Si:',sorted(set(SPECIES)-set(si_species_seen['Root silicon'])),flush=True)
SLIDES=[]

# Summary indices, all quantities dimensionless or percentage; upper percent changes descriptively, not necessarily significance.
def ratio(a,b):return a/b if np.isfinite(a) and np.isfinite(b) and b>0 else np.nan
def pct(a,b):return (a-b)/b*100 if np.isfinite(a) and np.isfinite(b) and b>0 else np.nan
indices=[]
for s,sp in enumerate(SPECIES):
    c,si,dr,ds=TOT[s,:]
    vi=TOT[s,:];cv=np.std(vi,ddof=1)/np.mean(vi) if np.mean(vi)>0 else np.nan
    stability=1/(1+cv) if np.isfinite(cv) else np.nan
    dr_ratio=ratio(dr,c);si_resp=pct(si,c);mit=pct(ds,dr)
    recovery=ratio(ds-dr,c-dr) if c>dr and dr>=0 else np.nan
    vals=dict(Species=sp,Control_Weight=c,Si_Weight=si,Drought_Weight=dr,Drought_Si_Weight=ds,
              Drought_Retention=dr_ratio,Si_Response_Pct=si_resp,Si_Mitigation_Pct=mit,
              Dry_Weight_Stability=stability,Across_Treatments_CV=cv,Recovery_Fraction=recovery,
              Drought_Si_GSW=MEANS[s,3,6],Drought_Si_Root_Silicon=MEANS[s,3,8],Drought_Si_Soil_Silicon=MEANS[s,3,9])
    indices.append(vals)
# scaled 0-100 within panel. Composite is exploratory equal ranking except weights transparent.
def percentile_score(vals):
    a=np.asarray(vals,float);good=np.isfinite(a);res=np.full(len(a),50.)
    if good.sum()>1:
        ranks=stats.rankdata(a[good],method='average')
        res[good]=100*(ranks-1)/(good.sum()-1)
    return res
for idx,weights in enumerate(indices):pass
p_mit=percentile_score([r['Si_Mitigation_Pct'] for r in indices]);p_ret=percentile_score([r['Drought_Retention'] for r in indices]);p_si=percentile_score([r['Drought_Si_Root_Silicon'] for r in indices]);p_stab=percentile_score([r['Dry_Weight_Stability'] for r in indices])
for i,r in enumerate(indices):r['Composite_Score_0_100']=.35*p_mit[i]+.25*p_ret[i]+.20*p_si[i]+.20*p_stab[i]
ranked=sorted(indices,key=lambda r:r['Composite_Score_0_100'],reverse=True)
for i,r in enumerate(ranked):r['Composite_Rank']=i+1
IK=list(ranked[0].keys())
write_csv('02_species_stress_silicon_indices_and_rankings.csv',IK,[[r.get(k,np.nan) for k in IK] for r in ranked])

# statistics: full three-way factorial, preserving separate biological observations per measurement.
anova_rows=[];anova_output={}
for k,name in enumerate(NAMES):
    rows=[(sp,tr,TREAT.index(tr) in [1,3],TREAT.index(tr) in [2,3],float(v)) for s,sp in enumerate(SPECIES) for t,tr in enumerate(TREAT) for v in D[s,t,k] if np.isfinite(v)]
    try:
        m=ols('y ~ C(Species) * C(Si) * C(Drought)',data={'Species':np.array([x[0] for x in rows]),'Si':np.array([x[2] for x in rows]),'Drought':np.array([x[3] for x in rows]),'y':np.array([x[4] for x in rows])}).fit()
        an=anova_lm(m,typ=2)
        error=float(an.loc['Residual','sum_sq'])
        for effect in [q for q in an.index if q!='Residual']:
            a=an.loc[effect];partial=float(a['sum_sq'])/(float(a['sum_sq'])+error) if (float(a['sum_sq'])+error)>0 else np.nan
            anova_rows.append([name,effect,int(len(rows)),float(a['df']),float(a['F']),float(a['PR(>F)']),partial])
        anova_output[name]=[(x, float(an.loc[x,'PR(>F)'])) for x in an.index if x!='Residual']
        print('ANOVA',name,'n',len(rows),flush=True)
    except Exception as ex:
        qc.append(('', '',name,'',np.nan,'ANOVA failed: '+str(ex)))
write_csv('03_factorial_ANOVA_species_Si_drought.csv',['Parameter','Effect','Observation_count','Degrees_freedom','F','p_value','Partial_eta_squared'],anova_rows)
# Tukey within each species, 4 treatments (not on fabricated combined dry biomass).
tukey_rows=[]
for k,name in enumerate(NAMES):
 for s,sp in enumerate(SPECIES):
  arrays=[D[s,t,k,np.isfinite(D[s,t,k])] for t in range(4)]
  if any(len(a)<2 for a in arrays):continue
  try:
   tuk=tukey_hsd(*arrays)
   for i,j,label in [(2,3,'Drought_vs_Drought_Si'),(0,1,'Control_vs_Silicon'),(0,2,'Control_vs_Drought'),(1,3,'Silicon_vs_Drought_Si')]:
    tukey_rows.append([name,sp,label,TREAT[i],TREAT[j],len(arrays[i]),len(arrays[j]),np.mean(arrays[j])-np.mean(arrays[i]),pct(np.mean(arrays[j]),np.mean(arrays[i])),float(tuk.pvalue[i,j]),bool(tuk.pvalue[i,j]<.05)])
  except Exception as ex:qc.append((sp,'',name,'',np.nan,'Tukey failed: '+str(ex)))
write_csv('04_tukey_species_specific.csv',['Parameter','Species','Contrast','Group_A','Group_B','n_A','n_B','B_minus_A','Percent_change_vs_A','Tukey_adjusted_p','Significant_0_05'],tukey_rows)
write_csv('00_data_quality_flags.csv',['Species','Treatment','Parameter','Replicate','Raw_value','Flag'],qc)

# Pearson and Spearman matrices on 20 species means in each treatment; FDR adjusted among correlations.
pearson_rows=[];spearman_rows=[];partial_rows=[]
for t,tr in enumerate(TREAT):
    A=MEANS[:,t,:]
    for i in range(10):
     for j in range(i+1,10):
      good=np.isfinite(A[:,i])&np.isfinite(A[:,j]);u=A[good,i];v=A[good,j]
      if len(u)<4 or np.std(u)==0 or np.std(v)==0:continue
      a,b=stats.pearsonr(u,v);c,d=stats.spearmanr(u,v)
      pearson_rows.append([tr,NAMES[i],NAMES[j],len(u),a,b]);spearman_rows.append([tr,NAMES[i],NAMES[j],len(u),c,d])
    # partial corr between root silicon and total dry weight conditioning on soil silicon
    v=np.stack([MEANS[:,t,8],TOT[:,t],MEANS[:,t,9]],axis=1);v=v[np.isfinite(v).all(axis=1)]
    if len(v)>4:
      X=np.c_[np.ones(len(v)),v[:,2]]
      ux=v[:,0]-X@np.linalg.lstsq(X,v[:,0],rcond=None)[0]
      vy=v[:,1]-X@np.linalg.lstsq(X,v[:,1],rcond=None)[0]
      r=float(np.corrcoef(ux,vy)[0,1]); df=len(v)-3
      p=2*stats.t.sf(abs(r)*np.sqrt(df/(1-r*r)),df) if abs(r)<1 else 0.
      partial_rows.append([tr,len(v),r,p])

def adjust_bh(rows):
    # FDR corrected within each treatment, not across pooled tests.
    new=[]
    for tr in TREAT:
       group=[r for r in rows if r[0]==tr]
       pp=np.array([r[-1] for r in group]);n=len(pp)
       order=np.argsort(pp);sortedadj=np.minimum.accumulate((pp[order]*n/np.arange(1,n+1))[::-1])[::-1]
       q=np.empty(n);q[order]=np.clip(sortedadj,0,1)
       new.extend([r+[q[idx]] for idx,r in enumerate(group)])
    return new
pearson_rows=adjust_bh(pearson_rows);spearman_rows=adjust_bh(spearman_rows)
write_csv('05_pearson_correlations_FDR.csv',['Treatment','Trait_A','Trait_B','n_species','r','raw_p','BH_FDR_q'],pearson_rows)
write_csv('06_spearman_correlations_FDR.csv',['Treatment','Trait_A','Trait_B','n_species','rho','raw_p','BH_FDR_q'],spearman_rows)
write_csv('07_partial_correlation_rootSi_totalDry_control_soilSi.csv',['Treatment','n_species','partial_r','p_value'],partial_rows)

# Canonical Correlation: 1 soil variable => ONE canonical component (not multiple).
# Small sample and potentially in-sample inflation: report 299 permutation refits for each group.
cca_rows=[];cca_weights=[];cca_viz={}
for t,tr in enumerate(TREAT):
    # limit plant block to four interpretable predictors, avoid redundant dry sum constituents.
    X=np.column_stack([MEANS[:,t,0],MEANS[:,t,1],TOT[:,t],MEANS[:,t,6],MEANS[:,t,8]])
    Y=MEANS[:,t,9].reshape(-1,1)
    keep=np.isfinite(X).all(axis=1)&np.isfinite(Y[:,0]);XX=X[keep];YY=Y[keep]
    sx=StandardScaler();sy=StandardScaler();XX=sx.fit_transform(XX);YY=sy.fit_transform(YY)
    if len(XX)<8:continue
    cc=CCA(n_components=1,max_iter=2000);xc,yc=cc.fit_transform(XX,YY)
    r=float(abs(np.corrcoef(xc[:,0],yc[:,0])[0,1]));perms=[]
    for b in range(299):
        Yp=YY[rng.permutation(len(YY))]
        z=CCA(n_components=1,max_iter=1000);xp,yp=z.fit_transform(XX,Yp)
        perms.append(abs(np.corrcoef(xp[:,0],yp[:,0])[0,1]))
    p=(1+np.sum(np.array(perms)>=r))/(len(perms)+1)
    cca_rows.append([tr,len(XX),1,r,p,np.quantile(perms,.95)])
    cca_viz[tr]=(r,p,np.quantile(perms,.95))
    for trait,weight in zip(['Shoot length','Root length','Total dry weight','Stomatal conductance','Root silicon'],cc.x_weights_[:,0]):cca_weights.append([tr,trait,float(weight)])
write_csv('08_CCA_permutation_plant_traits_vs_soil_silicon.csv',['Treatment','Complete_species_n','Max_possible_Canonical_Dimensions','Absolute_Canonical_r_insample','Permutation_p_299','Permutation_null_95_percentile'],cca_rows)
write_csv('09_CCA_plant_variable_weights.csv',['Treatment','Plant_variable','Weight_CC1'],cca_weights)

# ML at species-treatment mean level; group by species; avoid dry weight components in predictors.
X=np.column_stack([MEANS[:,:,k].reshape(-1) for k in [0,1,2,3,6,7,8,9]]+[np.tile((np.arange(4)==t).astype(float),20) for t in range(4)])
PREDICTORS=['Shoot length','Root length','Fresh shoot weight','Fresh root weight','Stomatal conductance','Leaf VPD','Root silicon','Soil silicon','Control indicator','Silicon indicator','Drought indicator','Drought+Si indicator']
y=TOT.reshape(-1);groups=np.repeat(np.array(SPECIES),4)
good=np.isfinite(y);X=X[good];y=y[good];groups=groups[good];sample_treat=np.tile(np.array(TREAT),20)[good]
models={
'Ridge':Ridge(alpha=10.),'Lasso':Lasso(alpha=2.,max_iter=25000,random_state=42),
'Elastic Net':ElasticNet(alpha=2.,l1_ratio=.5,max_iter=25000,random_state=42),
'Random Forest':RandomForestRegressor(n_estimators=180,min_samples_leaf=2,max_features=.85,random_state=42,n_jobs=1),
'XGBoost':XGBRegressor(n_estimators=180,max_depth=2,learning_rate=.04,subsample=.9,colsample_bytree=.9,reg_lambda=5.,objective='reg:squarederror',random_state=42,n_jobs=1),
'Gradient Boosting':GradientBoostingRegressor(n_estimators=120,max_depth=2,learning_rate=.05,min_samples_leaf=3,random_state=42),
'Support Vector Regression':SVR(kernel='rbf',C=100.,epsilon=5.)}
cv=GroupKFold(n_splits=5);random_cv=KFold(n_splits=5,shuffle=True,random_state=42)
train_ix,test_ix=next(GroupShuffleSplit(n_splits=1,test_size=.2,random_state=42).split(X,y,groups))
mlrows=[];foldrows=[];fits={};permodel={}
for nm,est in models.items():
    pipe=Pipeline([('imputer',SimpleImputer(strategy='median',add_indicator=False)),('scale',StandardScaler()),('model',est)])
    foldtrain=[];foldtest=[];foldmae=[];foldrmse=[]
    for fold,(tr,te) in enumerate(cv.split(X,y,groups),1):
        p=clone(pipe).fit(X[tr],y[tr]); pred=p.predict(X[te]); trainpred=p.predict(X[tr])
        foldtrain.append(r2_score(y[tr],trainpred));foldtest.append(r2_score(y[te],pred));foldmae.append(mean_absolute_error(y[te],pred));foldrmse.append(np.sqrt(mean_squared_error(y[te],pred)))
        foldrows.append([nm,'Species_grouped',fold,len(tr),len(te),','.join(sorted(set(groups[te]))),foldtrain[-1],foldtest[-1],foldmae[-1],foldrmse[-1]])
    rr=[]
    for fold,(tr,te) in enumerate(random_cv.split(X,y),1):
        p=clone(pipe).fit(X[tr],y[tr]);rr.append(r2_score(y[te],p.predict(X[te])))
        foldrows.append([nm,'Random_rows_NOT_primary',fold,len(tr),len(te),'Rows mixed by species',r2_score(y[tr],p.predict(X[tr])),rr[-1],mean_absolute_error(y[te],p.predict(X[te])),np.sqrt(mean_squared_error(y[te],p.predict(X[te])))])
    fitted=clone(pipe).fit(X[train_ix],y[train_ix]);pred=fitted.predict(X[test_ix]);fits[nm]=fitted
    mlrows.append([nm,float(np.mean(foldtrain)),float(np.std(foldtrain,ddof=1)),float(r2_score(y[train_ix],fitted.predict(X[train_ix]))),float(r2_score(y[test_ix],pred)),float(np.mean(foldtest)),float(np.std(foldtest,ddof=1)),float(np.mean(rr)),float(np.std(rr,ddof=1)),float(np.mean(foldmae)),float(np.mean(foldrmse)),int(len(X)),int(len(np.unique(groups)))])
    permodel[nm]=dict(grouped=np.array(foldtest),train=np.array(foldtrain),random=np.array(rr),hold=r2_score(y[test_ix],pred))
    print('ML',nm,'species-grouped CV R2',round(np.mean(foldtest),3),'±',round(np.std(foldtest,ddof=1),3),flush=True)
mlrows.sort(key=lambda z:z[5],reverse=True)
write_csv('10_ML_seven_models_grouped_5fold.csv',['Model','Train_fold_mean_R2','Train_fold_SD_R2','Holdout_train_R2','Holdout_test_R2_single','Species_grouped_CV_mean_R2','Species_grouped_CV_SD_R2','Random_row_CV_mean_R2_sensitivity_only','Random_row_CV_SD_R2','Grouped_CV_MAE_raw_units','Grouped_CV_RMSE_raw_units','n_species_treatment_rows','n_species'],mlrows)
write_csv('11_ML_fold_by_fold.csv',['Model','CV_strategy','Fold','Train_rows','Test_rows','Heldout_species','Train_R2','Test_R2','Test_MAE','Test_RMSE'],foldrows)
BEST=mlrows[0][0]; BEST_EST=models[BEST];pipe=Pipeline([('imputer',SimpleImputer(strategy='median')),('scale',StandardScaler()),('model',BEST_EST)])
# Out-of-fold predictions on species unseen during training.
oof=np.full(len(y),np.nan)
for tr,te in cv.split(X,y,groups):
 p=clone(pipe).fit(X[tr],y[tr]);oof[te]=p.predict(X[te])
write_csv('12_best_model_species_grouped_OOF_predictions.csv',['Species','Treatment','Actual_Total_Dry_Weight','Predicted_Total_Dry_Weight','Residual'],[[groups[i],sample_treat[i],y[i],oof[i],y[i]-oof[i]] for i in range(len(y))])
# Permutation importance averaged over independent species-held-out validation folds.
impvec=[]
for tr,te in cv.split(X,y,groups):
 p=clone(pipe).fit(X[tr],y[tr]);res=permutation_importance(p,X[te],y[te],scoring='r2',n_repeats=10,random_state=SEED,n_jobs=1)
 impvec.append(res.importances_mean)
impvec=np.array(impvec);impmeans=np.mean(impvec,axis=0);impsd=np.std(impvec,axis=0,ddof=1)
write_csv('13_best_model_species_grouped_permutation_importance.csv',['Feature','Mean_drop_in_R2_across_folds','SD_across_folds','Model'],[[p,a,b,BEST] for p,a,b in sorted(zip(PREDICTORS,impmeans,impsd),key=lambda z:z[1],reverse=True)])
# PCA on means, imputing median strictly for exploratory visualization (not applied before CV in ML).
feat=np.column_stack([MEANS[:,:,k].reshape(-1) for k in range(10)])
fimp=SimpleImputer(strategy='median').fit_transform(feat)
z=StandardScaler().fit_transform(fimp);pca=PCA().fit(z);scores=pca.transform(z)
write_csv('14_PCA_scores.csv',['Species','Treatment','PC1','PC2','PC3'],[[SPECIES[i//4],TREAT[i%4],*scores[i,:3]] for i in range(80)])
write_csv('15_PCA_loadings.csv',['Trait','PC1','PC2','PC3'],[[NAMES[k],*pca.components_[:3,k]] for k in range(10)])
write_csv('16_PCA_explained_variance.csv',['PC','Explained_ratio'],[[i+1,v] for i,v in enumerate(pca.explained_variance_ratio_)])
print('All statistical calculations done. Creating 600 dpi figures...',flush=True)

# Figure A summary of design and sample completeness
fig,axes=plt.subplots(1,2,figsize=(11,5))
img=axes[0].imshow(CNT.sum(axis=0).T,cmap='BuGn',aspect='auto')
axes[0].set_xticks(range(4),['C','Si','D','D+Si']);axes[0].set_yticks(range(10),NAMES)
for j in range(10):
 for t in range(4):axes[0].text(t,j,str(CNT[:,t,j].sum()),ha='center',va='center',color='#172a3a',fontsize=9)
axes[0].set_title('Valid measurements per treatment')
# normalized control (among species) median change in 8 plant traits
ch=np.array([[pct(MEANS[s,t,k],MEANS[s,0,k]) for s in range(20)] for k in range(8) for t in (1,2,3)]).reshape(8,3,20)
med=np.nanmedian(ch,axis=2)
heat=axes[1].imshow(med,cmap='RdBu_r',norm=TwoSlopeNorm(vcenter=0,vmin=min(-65,np.nanmin(med)),vmax=max(65,np.nanmax(med))),aspect='auto')
axes[1].set_yticks(range(8),NAMES[:8]);axes[1].set_xticks(range(3),['Si vs C','D vs C','D+Si vs C'])
for i in range(8):
 for j in range(3):axes[1].text(j,i,f'{med[i,j]:+.0f}%',ha='center',va='center',fontsize=8)
axes[1].set_title('Median within-species change vs control')
fig.colorbar(heat,ax=axes[1],fraction=.04,label='% change')
fig.suptitle('Data completeness and treatment response overview',fontsize=15,weight='bold');fig.tight_layout()
save_fig('01_study_overview',fig,'Study structure and treatment effects','Cell counts reflect available biological measurements; values are species-median percent changes.')

# Model bars
viz_names=[r[0] for r in mlrows if float(r[5])>=0]
fig,ax=plt.subplots(figsize=(11,5.5));names=viz_names;xv=np.arange(len(names));wbar=.26
ax.bar(xv-wbar,[permodel[m]['train'].mean() for m in names],wbar,color='#8396B4',label='Training folds',yerr=[permodel[m]['train'].std(ddof=1) for m in names],capsize=3)
ax.bar(xv,[permodel[m]['hold'] for m in names],wbar,color='#38A89D',label='Species holdout (single split)')
ax.bar(xv+wbar,[permodel[m]['grouped'].mean() for m in names],wbar,color='#D99B55',label='5-fold grouped CV',yerr=[permodel[m]['grouped'].std(ddof=1) for m in names],capsize=3)
ax.set_xticks(xv,names,rotation=24,ha='right');ax.set_ylabel('R² (unseen species for test/CV)');ax.set_ylim(bottom=0);ax.legend(fontsize=9,ncol=3,loc='lower left');ax.set_title('Six machine-learning models | dry biomass prediction');fig.tight_layout()
save_fig('02_ml_training_test_cv',fig,'Model performance: training, holdout test, and species-grouped 5-fold CV','Only models with non-negative grouped-CV mean R² are shown; CV bars show mean ± SD over five folds.',w=11,h=5.5)

fig,ax=plt.subplots(figsize=(11,5));s=[permodel[m]['grouped'] for m in names];parts=ax.boxplot(s,patch_artist=True,tick_labels=names)
for p in parts['boxes']:p.set_facecolor('#81C3BA');p.set_alpha(.75)
for i,ss in enumerate(s):ax.scatter(np.full(len(ss),i+1)+rng.normal(0,.045,len(ss)),ss,color='#205A64',s=12,zorder=4)
ax.set_ylim(bottom=0);ax.set_ylabel('Held-out species R² per fold');ax.set_xticklabels(names,rotation=23,ha='right');ax.set_title('Distribution of species-grouped five-fold scores');fig.tight_layout()
save_fig('03_ml_cv_boxplots',fig,'Distribution of model R² across five species-grouped folds','Every point represents a held-out fold containing four plant species; only the six non-negative grouped-CV models are displayed.')

fig,ax=plt.subplots(figsize=(11,5));xx=np.arange(len(names))
ax.bar(xx-.19,[permodel[m]['random'].mean() for m in names],.38,label='Random-row CV (optimistic)',color='#9FBDD4',yerr=[permodel[m]['random'].std(ddof=1) for m in names],capsize=3)
ax.bar(xx+.19,[permodel[m]['grouped'].mean() for m in names],.38,label='Grouped by species (primary)',color='#20877F',yerr=[permodel[m]['grouped'].std(ddof=1) for m in names],capsize=3)
ax.set_xticks(xx,names,rotation=25,ha='right');ax.set_ylim(bottom=0);ax.set_ylabel('5-fold mean test R² ± SD');ax.set_title('Generalization check: random rows vs unseen species');ax.legend();fig.tight_layout()
save_fig('04_grouped_vs_random_cv',fig,'Sensitivity of model results to data-splitting design','Primary selection uses species-grouped CV, not the potentially optimistic row-random CV; six non-negative grouped-CV models are shown.')

fig,ax=plt.subplots(figsize=(8,5.7));
for t,tr in enumerate(TREAT):
 ii=np.where(sample_treat==tr)[0];ax.scatter(y[ii],oof[ii],s=38,alpha=.75,label=tr,color=COLORS[t])
lim=[min(y.min(),oof.min())-.07*np.ptp(y),max(y.max(),oof.max())+.07*np.ptp(y)]
ax.plot(lim,lim,'--',color='#737373');ax.set(xlim=lim,ylim=lim,xlabel='Observed total dry weight (original numeric scale)',ylabel='Out-of-fold predicted total dry weight');ax.set_title(f'Best grouped-CV model: {BEST}\nOOF R² = {r2_score(y,oof):.3f} | n = {len(y)}');ax.legend(fontsize=8,loc='upper left');fig.tight_layout()
save_fig('05_best_model_predicted_vs_observed',fig,'Held-out species predictions from the best model','Every species-treatment point is predicted by a model trained without that species.',w=8,h=5.7)

fig,ax=plt.subplots(figsize=(9.3,5.5));keep=np.where(impmeans>=0)[0];ord_keep=keep[np.argsort(impmeans[keep])]
ax.barh(np.array(PREDICTORS)[ord_keep],impmeans[ord_keep],xerr=impsd[ord_keep],color='#DBA66B',capsize=3)
ax.set_xlim(left=0);ax.set_xlabel('Mean decrease in held-out R² ± SD across five folds');ax.set_title(f'Permutation importance | {BEST}');fig.tight_layout()
save_fig('06_permutation_importance',fig,'Fold-averaged permutation feature importance','Only non-negative mean permutation-importance features are displayed; importance describes predictive sensitivity, not causality.',w=9.3,h=5.5)

# four Pearson matrix heatmaps; trait subsets include soil and root silicon
for t,tr in enumerate(TREAT):
 a=MEANS[:,t,:];corr=np.full((10,10),np.nan)
 for i in range(10):
  for j in range(10):
   ok=np.isfinite(a[:,i])&np.isfinite(a[:,j]);corr[i,j]=np.corrcoef(a[ok,i],a[ok,j])[0,1] if ok.sum()>3 and np.std(a[ok,i])>0 and np.std(a[ok,j])>0 else np.nan
 fig,ax=plt.subplots(figsize=(9,7));im=ax.imshow(corr,cmap='RdBu_r',vmin=-1,vmax=1)
 ax.set_xticks(range(10),NAMES,rotation=55,ha='right',fontsize=8);ax.set_yticks(range(10),NAMES,fontsize=9)
 for i in range(10):
  for j in range(10):
   if np.isfinite(corr[i,j]):ax.text(j,i,f'{corr[i,j]:+.2f}',ha='center',va='center',fontsize=7,color='white' if abs(corr[i,j])>.7 else '#2c2c2c')
 fig.colorbar(im,ax=ax,fraction=.046,pad=.04,label='Pearson r')
 ax.set_title(f'Pearson correlations | {tr} | n = 19–20 species');fig.tight_layout()
 save_fig(f'{7+t:02d}_pearson_{t+1}',fig,f'Pearson correlation matrix — {tr}','Species-level treatment means; n=19 when silicon variable missing in Bell Pepper.',w=9,h=7)

fig,ax=plt.subplots(figsize=(8.6,5.4));ids=[r[0] for r in cca_rows];vals=np.array([r[3] for r in cca_rows]);pbs=np.array([r[4] for r in cca_rows]);nulls=np.array([r[5] for r in cca_rows]);zz=np.arange(len(ids))
ax.bar(zz,vals,color=COLORS,width=.57)
for i,v in enumerate(vals):ax.text(i,v+.015,f'r={v:.2f}\np={pbs[i]:.3f}',ha='center',va='bottom',fontsize=10)
ax.plot(zz,nulls,'k--o',label='95th percentile of permutation null')
ax.set_xticks(zz,['Control','Silicon','Drought','Drought+Si']);ax.set_ylim(0,max(1.15,max(vals)+.17));ax.set_ylabel('|Canonical correlation|');ax.set_title('CCA: five plant traits versus soil silicon');ax.legend(fontsize=8);fig.tight_layout()
save_fig('11_cca_plant_soil',fig,'Canonical correlation analysis under four treatments','Only one canonical component is possible because there is one soil variable; permutation tests refit the CCA 299 times.',w=8.6,h=5.4)

fig,axes=plt.subplots(2,2,figsize=(10,8));
for t,tr in enumerate(TREAT):
 ax=axes.flat[t];x0=MEANS[:,t,9];y0=MEANS[:,t,8];ok=np.isfinite(x0)&np.isfinite(y0)
 ax.scatter(x0[ok],y0[ok],s=36,alpha=.8,color=COLORS[t]);r,p=stats.pearsonr(x0[ok],y0[ok])
 if ok.sum()>3:
  slope,inter=np.polyfit(x0[ok],y0[ok],1);xp=np.linspace(x0[ok].min(),x0[ok].max(),80);ax.plot(xp,slope*xp+inter,color='#444',ls='--',lw=1)
 ax.set_title(f'{tr}: r={r:+.2f}, p={p:.3g}, n={ok.sum()}');ax.set(xlabel='Soil silicon (%)',ylabel='Root silicon (%)')
fig.suptitle('Root and soil Si relationships across species',weight='bold');fig.tight_layout()
save_fig('12_root_soil_Si_associations',fig,'Root versus soil silicon concentrations across treatments','Across-species associations do not establish causal uptake relationships.',w=10,h=8)

fig,ax=plt.subplots(figsize=(9.5,6));
for t,tr in enumerate(TREAT):
 ii=np.arange(t,80,4);ax.scatter(scores[ii,0],scores[ii,1],c=COLORS[t],label=tr,s=40,alpha=.76)
for i in np.argsort(np.linalg.norm(scores[:,:2],axis=1))[-7:]:ax.annotate(SPECIES[i//4],scores[i,:2],fontsize=7,xytext=(3,3),textcoords='offset points')
ax.axhline(0,color='#ccc');ax.axvline(0,color='#ccc');ax.set_xlabel(f'PC1 ({pca.explained_variance_ratio_[0]*100:.1f}% variance)');ax.set_ylabel(f'PC2 ({pca.explained_variance_ratio_[1]*100:.1f}% variance)');ax.legend(fontsize=8,ncol=2,loc='best');ax.set_title('PCA of ten measured traits | 80 species-treatment means');fig.tight_layout()
save_fig('13_pca_species_treatment',fig,'Principal component analysis across four treatments','Trait values z-scored; missing silicon means for Bell Pepper median-imputed for PCA only.',w=9.5,h=6)

fig,axes=plt.subplots(1,2,figsize=(10,4.8));axes[0].bar(np.arange(1,11),pca.explained_variance_ratio_*100,color='#53A9A0');axes[0].set_xlabel('Principal component');axes[0].set_ylabel('Explained variance (%)');axes[0].set_xticks(range(1,11));axes[0].set_title('PCA scree plot')
order=np.argsort(abs(pca.components_[0]));axes[1].barh(np.array(NAMES)[order],pca.components_[0,order],color=['#487D99' if v>=0 else '#D18E65' for v in pca.components_[0,order]]);axes[1].axvline(0,color='#333');axes[1].set_title('PC1 feature loadings');fig.tight_layout()
save_fig('14_pca_scree_loadings',fig,'PCA variance explained and trait contributions','Loadings indicate associations with PC1, not independent causal effects.',w=10,h=4.8)

# Total dry weight treatment heatmap per species relative to its own control
vals=np.stack([[pct(TOT[s,t],TOT[s,0]) for t in (1,2,3)] for s in range(20)])
fig,ax=plt.subplots(figsize=(8.8,7.2));mm=max(50,np.nanpercentile(abs(vals),95));im=ax.imshow(vals,cmap='RdBu_r',norm=TwoSlopeNorm(vcenter=0,vmin=-mm,vmax=mm),aspect='auto')
ax.set_yticks(range(20),SPECIES);ax.set_xticks(range(3),['Si vs C','D vs C','D+Si vs C'])
for s in range(20):
 for j in range(3):ax.text(j,s,f'{vals[s,j]:+.0f}%',ha='center',va='center',fontsize=8)
fig.colorbar(im,ax=ax,fraction=.035,label='% change in total dry weight');ax.set_title('Total dry weight response per species');fig.tight_layout()
save_fig('15_species_dry_biomass_heatmap',fig,'Species-specific dry-weight responses to silicon and drought','Total dry weight is sum of treatment-mean shoot and root dry weight; not a paired replicate-level sum.',w=8.8,h=7.2)

fig,ax=plt.subplots(figsize=(9.4,6.2))
xx=np.array([r['Drought_Retention'] for r in indices]);yy=np.array([r['Si_Mitigation_Pct'] for r in indices]);cc=np.array([r['Drought_Si_Root_Silicon'] for r in indices])
chosen=set([r['Species'] for r in ranked[:5]]+['Lettuce','Millet','Bell Pepper','White Radish'])
ok=np.isfinite(cc)
sc=ax.scatter(xx[ok],yy[ok],c=cc[ok],cmap='viridis',s=87,alpha=.9,edgecolors='white',linewidth=.7)
ax.scatter(xx[~ok],yy[~ok],marker='X',color='#6B7280',s=94,label='Root silicon not measured')
for i,sp in enumerate(SPECIES):
 if sp in chosen:
  offset=(6,5) if sp not in ['Corn','Soybean'] else (7,-13)
  ax.annotate(sp,(xx[i],yy[i]),xytext=offset,textcoords='offset points',fontsize=9)
ax.axhline(0,color='#999',ls='--');ax.axvline(1,color='#999',ls='--')
ax.set_xlabel('Drought dry-weight retention (D / Control)');ax.set_ylabel('Silicon mitigation: (D+Si − D) / D × 100 (%)')
ax.set_title('Drought tolerance versus silicon-mediated biomass gain')
fig.colorbar(sc,ax=ax,label='Measured root silicon under D+Si (%)');ax.legend(loc='upper left',fontsize=9);fig.tight_layout()
save_fig('16_drought_retention_vs_mitigation',fig,'Drought retention versus silicon-mediated dry-weight improvement','Missing Bell Pepper root silicon shown as an X; selected species annotated.',w=9.4,h=6.2)

fig,ax=plt.subplots(figsize=(9,7.1));scrows=ranked[::-1];ax.barh([r['Species'] for r in scrows],[r['Composite_Score_0_100'] for r in scrows],color=['#1A8A82' if r['Composite_Rank']<=5 else '#9FB6C4' for r in scrows]);ax.set_xlim(0,102);ax.set_xlabel('Exploratory composite score (0–100)');ax.set_title('Cross-species Si × drought performance ranking');fig.tight_layout()
save_fig('17_integrated_species_ranking',fig,'Composite ranking of all 20 species','Weighted percentile ranks: mitigation 35%, drought retention 25%, root Si 20%, across-treatment stability 20%.',w=9,h=7.1)

fig,ax=plt.subplots(figsize=(9.5,6.2))
sx=np.array([r['Dry_Weight_Stability'] for r in indices]);sy=np.array([r['Si_Mitigation_Pct'] for r in indices]);co=np.array([r['Drought_Si_GSW'] for r in indices]);bas=np.array([r['Si_Response_Pct'] for r in indices]);q=np.nanmedian(bas)
chosen=set([r['Species'] for r in ranked[:5]]+['Lettuce','Millet','Bell Pepper','White Radish'])
for marker,cat,sel in [('o','Si response ≥ median',bas>=q),('D','Si response < median',bas<q)]:
 p=ax.scatter(sx[sel],sy[sel],c=co[sel],cmap='viridis',vmin=np.nanmin(co),vmax=np.nanmax(co),s=99,marker=marker,edgecolors='white',linewidth=.8,label=cat)
for i,sp in enumerate(SPECIES):
 if sp in chosen:
  offset=(6,5) if sp not in ['Corn','Soybean','Millet'] else (8,-12)
  ax.annotate(sp,(sx[i],sy[i]),xytext=offset,textcoords='offset points',fontsize=9)
ax.axhline(0,color='#999',ls='--');ax.set_xlabel('Stability = 1 / (1 + CV across four dry-weight means)');ax.set_ylabel('Silicon mitigation under drought (%)');ax.set_title('Integrated silicon performance: stability × mitigation')
ax.legend(loc='upper left',fontsize=8);fig.colorbar(p,ax=ax,label='Stomatal conductance under D+Si (not photosynthesis)');fig.tight_layout()
save_fig('18_integrated_stability_SMI_gsw',fig,'Integrated species stability versus silicon mitigation','Color uses stomatal conductance (not photosynthesis); shapes encode above/below median Si response; selected species annotated.',w=9.5,h=6.2)

# treatment trends GSW and VPD using species normalized medians (avoid mean across incompatible crop scales)
fig,axs=plt.subplots(1,2,figsize=(10,4.7))
for ax,k in zip(axs,(6,7)):
 xx=np.array([[pct(MEANS[s,t,k],MEANS[s,0,k]) for s in range(20)] for t in range(4)])
 med=np.nanmedian(xx,axis=1);qlo=np.nanpercentile(xx,25,axis=1);qhi=np.nanpercentile(xx,75,axis=1)
 ax.errorbar(range(4),med,yerr=[med-qlo,qhi-med],fmt='o-',capsize=5,color=ACCENT);ax.axhline(0,color='#999',ls='--');ax.set_xticks(range(4),['C','Si','D','D+Si']);ax.set_title(NAMES[k]);ax.set_ylabel('Median change vs Control (%)\n(error: species IQR)')
fig.suptitle('Physiological response under silicon and drought',weight='bold');fig.tight_layout()
save_fig('19_gsw_vpd_treatment_response',fig,'Physiological responses: stomatal conductance and leaf VPD','Median paired across species; error bars are interquartile ranges among species.',w=10,h=4.7)

fig,axs=plt.subplots(1,2,figsize=(11,7),sharey=True)
for ax,k in zip(axs,(8,9)):
 mat=MEANS[:,3,k];ind=np.argsort(np.where(np.isfinite(MEANS[:,3,8]),MEANS[:,3,8],np.inf))
 ax.barh(np.array(SPECIES)[ind],mat[ind],color=ACCENT if k==8 else '#87AFC5');ax.set_xlabel(NAMES[k]+' (%)');ax.set_title('Drought + Si | '+NAMES[k]);ax.grid(axis='x',alpha=.15)
fig.suptitle('Silicon distribution across 20 species',weight='bold');fig.tight_layout()
save_fig('20_root_soil_Si_by_species',fig,'Root and soil silicon under silicon plus drought','Bell Pepper silicon values are missing, not zero. These are descriptive concentrations.',w=11,h=7)

# Stat summary heatmap -log10 p of main and interaction terms
terms=['C(Species)','C(Si)','C(Drought)','C(Si):C(Drought)','C(Species):C(Si)','C(Species):C(Drought)','C(Species):C(Si):C(Drought)']
termshort=['Species','Si','Drought','Si×Drought','Species×Si','Species×Drought','Species×Si×Drought']
heat=np.full((10,len(terms)),np.nan)
for k,name in enumerate(NAMES):
 for j,term in enumerate(terms):
  for row in anova_rows:
   if row[0]==name and row[1]==term:heat[k,j]=-np.log10(max(float(row[5]),1e-300))
fig,ax=plt.subplots(figsize=(10.5,6.9));c=ax.imshow(np.minimum(heat,12),cmap='YlGnBu',vmin=0,vmax=12,aspect='auto')
ax.set_xticks(range(7),termshort,rotation=38,ha='right');ax.set_yticks(range(10),NAMES)
for i in range(10):
 for j in range(7):
  if np.isfinite(heat[i,j]):ax.text(j,i,'***' if heat[i,j]>3 else '**' if heat[i,j]>2 else '*' if heat[i,j]>1.3 else 'ns',ha='center',va='center',color='white' if heat[i,j]>5 else '#263342',fontsize=9)
ax.set_title('Factorial ANOVA: species × Si × drought (per trait)');fig.colorbar(c,ax=ax,fraction=.04,label='−log10(p), clipped at 12');fig.tight_layout()
save_fig('21_ANOVA_significance_heatmap',fig,'Factorial ANOVA across ten measured traits','*** p<0.001; ** p<0.01; * p<0.05; ns otherwise. Tests use raw replicates per trait, not merged pseudo-replicates.',w=10.5,h=6.9)

# Species mitigation, shoot/root separate avoid assuming paired total biomass variances
fig,ax=plt.subplots(figsize=(9.2,7));r=np.arange(20);ind=np.argsort([x['Si_Mitigation_Pct'] for x in indices]);
barvals=vals[ind,:] if False else np.stack([[pct(MEANS[s,3,4],MEANS[s,2,4]),pct(MEANS[s,3,5],MEANS[s,2,5])] for s in ind]);
ax.barh(r-.19,barvals[:,0],.36,color='#329C95',label='Shoot dry weight');ax.barh(r+.19,barvals[:,1],.36,color='#E0AA71',label='Root dry weight');ax.set_yticks(r,np.array(SPECIES)[ind]);ax.axvline(0,color='#333');ax.set_xlabel('Drought + Si vs Drought: change (%)');ax.set_title('Biomass-specific silicon mitigation by species');ax.legend();fig.tight_layout()
save_fig('22_root_shoot_Si_mitigation',fig,'Shoot and root dry-weight mitigation under drought','Descriptive within-species percentage change; statistical contrasts provided separately in Tukey CSV.',w=9.2,h=7)

# Key winner tables + narratives
bestrow=mlrows[0];main_rank=[(r['Species'],r['Composite_Score_0_100']) for r in ranked[:10]]
importances=sorted(zip(PREDICTORS,impmeans),key=lambda z:z[1],reverse=True)[:6]
summary={
'input_workbook':INPUT.name,'species_count':20,'treatments':TREAT,'variables':NAMES,
'valid_observations_per_parameter':{NAMES[k]:int(CNT[:,:,k].sum()) for k in range(10)},
'flags':qc,'available_silicon_species_count':len(si_species_seen['Root silicon']),
'best_model_by_5fold_species_grouped_cv':BEST,
'best_model_grouped_CV_mean_R2':bestrow[5],'best_model_grouped_CV_SD_R2':bestrow[6],
'best_model_species_holdout_test_R2_single_split':bestrow[4],
'best_model_grouped_CV_MAE':bestrow[9],
'best_model_grouped_CV_RMSE':bestrow[10],
'best_model_out_of_fold_R2_pooled':float(r2_score(y,oof)),
'best_model_top_permutation_features':importances,
'exploratory_top_10_species':main_rank,
'highest_biomass_mitigation_species':max(indices,key=lambda r:r['Si_Mitigation_Pct'])['Species'],
'highest_drought_retention_species':max(indices,key=lambda r:r['Drought_Retention'])['Species'],
'highest_rootSi_under_Si_drought_species':max((r for r in indices if np.isfinite(r['Drought_Si_Root_Silicon'])),key=lambda r:r['Drought_Si_Root_Silicon'])['Species'],
'cautions':['Biomass worksheet dry-weight units not explicitly provided; graph uses original numerical units.','Photosynthesis was NOT measured; GSW substitutes only for chart point color, not physiological photosynthetic rate.','One negative Okra root silicon measurement excluded, retained in raw-audit CSV.','Bell Pepper root and soil silicon missing for all treatments.','CCA has one soil variable, hence at most one canonical correlation per treatment. In-sample CCA may overfit with 19 species.','No plant replicate IDs link the ten sheets. Individual parameters are analyzed on their own biological replicates; multivariate models use species x treatment means (n=80).','Composite ranking is exploratory and sensitive to chosen metrics and weights.','Machine learning is predictive, not evidence of causal silicon mechanisms; grouped-CV generalization may be weak or negative.']}
(OUT/'00_analysis_summary.json').write_text(json.dumps(summary,indent=2,default=lambda x:float(x) if isinstance(x,(np.floating,np.integer)) else str(x)),encoding='utf8')

# Professional PPT: figure-only white slides, subtitle and method note on each.
prs=Presentation();prs.slide_width=Inches(13.333);prs.slide_height=Inches(7.5)
blank=prs.slide_layouts[6]
def textbox(sl,x,y,w,h,text,size,color,bold=False):
 tb=sl.shapes.add_textbox(Inches(x),Inches(y),Inches(w),Inches(h));tf=tb.text_frame;tf.word_wrap=True
 for idx,line in enumerate(text.split('\n')):
  p=tf.paragraphs[0] if idx==0 else tf.add_paragraph();p.text=line;p.font.name='Aptos';p.font.size=Pt(size);p.font.bold=bold;p.font.color.rgb=RGBColor(*color)
 return tb

def slide_base(title,subtitle='',note=''):
 sl=prs.slides.add_slide(blank)
 stripe=sl.shapes.add_shape(MSO_SHAPE.RECTANGLE,0,0,prs.slide_width,Inches(.15));stripe.fill.solid();stripe.fill.fore_color.rgb=RGBColor(20,125,123);stripe.line.fill.background()
 textbox(sl,.55,.3,12.3,.4,title,20,(27,46,66),True)
 if subtitle:textbox(sl,.55,.76,12.1,.38,subtitle,10,(77,94,109))
 if note:textbox(sl,.6,7.09,11.9,.3,note,8,(92,102,112))
 textbox(sl,12.08,7.07,.7,.25,f'{len(prs.slides):02d}',8,(110,119,130))
 return sl

sl=slide_base('Silicon × drought: integrated comparative plant analysis','20 species · 4 treatments · 10 parameters · 7 ML algorithms · species-grouped 5-fold CV','Source: supplied Excel workbook | Python analysis | 600 dpi figures')
textbox(sl,1.0,1.7,11.5,1.35,'Multivariate phenotypes, silicon mobilization, drought response, and predictive modeling',30,(26,60,75),True)
textbox(sl,1.0,3.45,10.8,1.4,'Study levels: 80 species × treatment means for ML/correlation, and raw biological replicates for factorial ANOVA and Tukey tests.\nAll ranks and predictions are exploratory and should be validated independently.',17,(65,87,102))
sl=slide_base('Key analytical findings',f'Species-grouped 5-fold model selection | best model: {BEST}','Overall ranks are index-dependent; silicon measurements are missing for Bell Pepper.')
msg=(f"Best CV mean R²: {bestrow[5]:+.3f} ± {bestrow[6]:.3f}     |     holdout test R²: {bestrow[4]:+.3f}\n"
     f"Best exploratory composite species: {ranked[0]['Species']} (score {ranked[0]['Composite_Score_0_100']:.1f}/100)\n"
     f"Largest relative dry-weight silicon mitigation: {summary['highest_biomass_mitigation_species']}\n"
     f"Highest drought dry-weight retention: {summary['highest_drought_retention_species']}\n"
     f"Highest root silicon under Si + drought: {summary['highest_rootSi_under_Si_drought_species']}")
textbox(sl,.8,1.65,11.8,4.4,msg,18,(35,60,76))
for title,caption,note in SLIDES:
 sl=slide_base(caption,subtitle='',note=note)
 path=str(FIG/(title+'.png'))
 from PIL import Image
 with Image.open(path) as im: iw,ih=im.size
 maxw,maxh=12.45,5.77
 scale=min(maxw/iw,maxh/ih)
 dw,dh=iw*scale,ih*scale
 sl.shapes.add_picture(path,Inches((13.333-dw)/2),Inches(1.14+(maxh-dh)/2),width=Inches(dw),height=Inches(dh))
sl=slide_base('Interpretation and limitations','Critical methodological safeguards before manuscript submission','See associated analysis tables and source code for all exact computations.')
points=['No photosynthesis parameter is present: stomatal conductance is NOT equivalent to assimilation.',
        'Plant and soil CCA can yield only one canonical dimension, because soil silicon is the only soil measure.',
        'Bell Pepper lacks both root/soil silicon; one negative Okra root silicon observation was excluded.',
        'Across-trait biological replicate IDs are unavailable; never treat arbitrary replicate rows as matched plants.',
        'Species-held-out grouped CV is the primary indicator of transferability; random-row CV is only a sensitivity check.',
        'Dry biomass units not specified in Excel; total is a sum of treatment means in original recorded units.',
        'Species stability and composite rankings are descriptive rather than validated genetic selection indices.']
for i,line in enumerate(points):textbox(sl,.85,1.2+i*.7,11.8,.59,'• '+line,13,(49,72,84))
prs.save(str(OUT/'Silicon_Drought_ML_and_Statistics_600dpi.pptx'))
# README / figure notes
with open(OUT/'README_Methods_and_Results.txt','w',encoding='utf8') as f:
 f.write('SILICON x DROUGHT | 20 PLANT SPECIES | MACHINE LEARNING AND STATISTICS\n'+'='*78+'\n\n')
 f.write('INPUT: '+str(INPUT.name)+'\n')
 f.write('LAYOUT: 10 sheets: shoot/root length n=10; fresh/dry shoot/root n=5; gs, VPD, root/soil Si n=3 per species/treatment.\n')
 f.write('ML: 80 species-treatment means; target=sum of mean shoot and root dry weights, as recorded (units unverified). Excludes dry shoot/root components from predictors. Uses species one-hot? NO: species identifiers are not predictors. Treatment is 4 one-hot indicators.\n')
 f.write('CV: GroupKFold(5) grouping all 4 treatments for each species together; also a fixed species holdout 20% and a random-row KFold sensitivity analysis.\n')
 f.write('All preprocessing (median imputation and scaling) fitted within each training fold only.\n')
 f.write('ANOVA: per-trait raw observations; fully crossed Species × Silicon × Drought, Type II sums of squares and partial eta squared.\n')
 f.write('Tukey: within-species 4 treatment comparisons for each original trait; no synthetic total dry weight replicates.\n')
 f.write('CCA: species treatment means; plant block (shoot length, root length, total dry weight, gs, root Si) vs soil Si. One canonical dimension maximum; 299 refit-permutation baseline; in-sample CC should be interpreted conservatively.\n')
 f.write('All Si and stress indices are calculated on species-treatment MEANS; composite score weights = 0.35 mitigation, 0.25 drought retention, 0.20 root Si D+Si, 0.20 dry-weight stability. Missing Si assigned neutral percentile 50 in composite.\n')
 f.write('Si Mitigation %=100*(D+Si − D)/D; dry-weight drought retention=D/Control; Si response %=100*(Si−Control)/Control; Stability=1/(1+CV across four dry-weight cell means), ddof=1.\n')
 f.write('NO measured photosynthesis: color of integrated chart uses D+Si stomatal conductance (not assimilation).\n')
 f.write('Bell Pepper soil/root Si missing in entire source; one negative Okra drought root Si replicate discarded from analytical values, preserved in audit.\n\n')
 f.write('BEST MODEL: '+BEST+'\nSpecies-grouped CV R² mean ± SD: '+f'{bestrow[5]:.4f} ± {bestrow[6]:.4f}'+'\nSingle species-holdout R²: '+f'{bestrow[4]:.4f}'+'\n')
 f.write('RANKING (EXPLORATORY):\n')
 for r in ranked:f.write(f"{r['Composite_Rank']:2d}. {r['Species']:20} {r['Composite_Score_0_100']:.2f}\n")
 f.write('\nFIGURES IN PPT:\n')
 for i,(name,caption,note) in enumerate(SLIDES,1):f.write(f'{i:02d} {name}: {caption}. {note}\n')
 f.write('\nREPRODUCTION: python Silicon_Drought_Analysis.py "Batch2-Silicate experiment values.xlsx" "Silicon_Drought_Results"\n')
 f.write('Packages: artifact_tool, numpy, scipy, scikit-learn, statsmodels, matplotlib, xgboost, python-pptx, pillow.\n')
print('SUCCESS',len(SLIDES),'figures; ppt slides',len(prs.slides),'best',BEST,'grouped mean',round(bestrow[5],3),flush=True)

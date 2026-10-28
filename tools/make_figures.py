import matplotlib; matplotlib.use('Agg')
import matplotlib.pyplot as plt, sys
plt.rcParams.update({'font.family':'Liberation Serif','font.size':11,'axes.spines.top':False,'axes.spines.right':False})
out=sys.argv[1]
S=['S2 Alternative\nmaintenance','S3 Major erosion\ncontrol','S4 Full\nbetterment']
col=['#7fa7c9','#3d7ab5','#1b3f66']
finNPV=[1251412.73,100837.73,-353372.91]; ben=[4323667.18,11595289.26,18522983.26]
ce=[-220.47,-6.62,14.53]; tpy=[189.2,507.45,810.55]
val=1721.111952; k=[b/val for b in ben]
# Fig1
fig,ax=plt.subplots(1,2,figsize=(9,3.6))
b=ax[0].bar(S,tpy,color=col); ax[0].set_ylabel('Fine sediment avoided (t/yr)'); ax[0].set_title('(a) Sediment avoided vs BAU',fontsize=11)
for r,v in zip(b,tpy): ax[0].text(r.get_x()+r.get_width()/2,v+10,f'{v:,.0f}',ha='center')
b=ax[1].bar(S,ce,color=col); ax[1].axhline(0,color='k',lw=.8); ax[1].set_ylabel('PV net financial cost ($/t abated)'); ax[1].set_title('(b) Cost-effectiveness',fontsize=11)
for r,v in zip(b,ce): ax[1].text(r.get_x()+r.get_width()/2,v+(6 if v>=0 else -18),f'${v:,.2f}' if v>=0 else f'−${-v:,.2f}',ha='center')
ax[1].set_ylim(-250,50)
for a in ax: a.tick_params(axis='x',labelsize=9.5)
fig.tight_layout(); fig.savefig(f'{out}/fig1_cea.png',dpi=200)
# Fig2
fig,ax=plt.subplots(figsize=(7,3.8)); x=range(3)
ax.bar(x,[f/1e6 for f in finNPV],color='#c98b3d',label='Financial NPV (net maintenance saving less capital)')
ax.bar(x,[b/1e6 for b in ben],bottom=[max(f,0)/1e6 for f in finNPV],color='#3d7ab5',label='PV of avoided sediment (abatement cost)')
tot=[f+b for f,b in zip(finNPV,ben)]
for i,t in enumerate(tot): ax.text(i,(max(finNPV[i],0)+ben[i])/1e6+0.4,f'NPV ${t/1e6:,.2f}M',ha='center',fontweight='bold')
ax.axhline(0,color='k',lw=.8); ax.set_xticks(list(x)); ax.set_xticklabels(S,fontsize=9.5); ax.set_ylabel('AUD million (PV, 7%, FY2026 $)'); ax.set_ylim(-1.5,21.5)
ax.legend(frameon=False,fontsize=9,loc='upper left'); fig.tight_layout(); fig.savefig(f'{out}/fig2_npv.png',dpi=200)
# Fig3
import numpy as np
v=np.linspace(0,3500,200); fig,ax=plt.subplots(figsize=(7,3.8))
for i in range(3): ax.plot(v,(finNPV[i]+k[i]*v)/1e6,color=col[i],lw=2,label=S[i].replace('\n',' '))
sw=(finNPV[0]-finNPV[2])/(k[2]-k[0]); print('switch S4>S2',sw,'S3>S2',(finNPV[0]-finNPV[1])/(k[1]-k[0]),'S4>S3',(finNPV[1]-finNPV[2])/(k[2]-k[1]))
ax.axvline(sw,color='grey',ls=':',lw=1); ax.text(sw+40,-4.5,f'S4 overtakes S2\nat ${sw:,.0f}/t',fontsize=9)
ax.axvline(val,color='k',ls='--',lw=1); ax.text(val+40,32,f'Base case\n${val:,.0f}/t',fontsize=9)
ax.set_ylim(-6,40); ax.axhline(0,color='k',lw=.6); ax.set_xlabel('Avoided abatement cost value ($/t, FY2026)'); ax.set_ylabel('NPV vs BAU (AUD million)')
ax.legend(frameon=False,fontsize=9,loc='upper left'); fig.tight_layout(); fig.savefig(f'{out}/fig3_sensitivity.png',dpi=200)

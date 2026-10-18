// Usage: NODE_PATH=$(npm root -g) node build_deck.js <data.json> <out.pptx>
// Message bars are partly data-driven (template strings); static wording (expert slide, action plan, notes) must be re-read after each refresh.
const pptxgen=require('pptxgenjs');const R=require(require('path').resolve(process.argv[2]||'/tmp/R2.json'));const OUTF=process.argv[3]||'Mahisagar_District_TB_Review.pptx';const _d=new Date(R.asof);const AS=_d.toLocaleDateString('en-GB',{day:'2-digit',month:'short',year:'numeric'}).replace(/ /g,'-');const _r=new Date(R.reg_asof);const RA=_r.toLocaleDateString('en-GB',{day:'2-digit',month:'short',year:'numeric'}).replace(/ /g,'-');
const p=new pptxgen();p.layout='LAYOUT_16x9';p.title='Mahisagar District TB Programme Performance Review 2026';
const NAVY='0B3C49',TEAL='1B7F8C',RED='D1495B',AMB='E9A23B',GRN='2E8B57',LIGHT='EEF4F5',INK='1F2A30',MUT='5F6F76';
const T=['Balasinor','Kadana','Khanpur','Lunawada','Santrampur','Virpur'];
const MS=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep'];
const pc=(a,b,d=0)=>b?(100*a/b).toFixed(d)+'%':'–';const n=x=>Number(x).toLocaleString('en-IN');
const rag=(v,hi,lo)=>v>=hi?'D5EEDC':v>=lo?'FCEBC4':'F6CFC8';
let sl=0;
function base(sec,title,msg){const s=p.addSlide();s.background={color:'FFFFFF'};sl++;
 s.addText(sec.toUpperCase(),{x:.4,y:.14,w:7,h:.22,fontFace:'Calibri',fontSize:9,bold:true,color:TEAL,margin:0,isTextBox:true});
 s.addText(title,{x:.4,y:.34,w:9.2,h:.5,fontFace:'Cambria',fontSize:22,bold:true,color:NAVY,margin:0,isTextBox:true});
 if(msg){s.addShape(p.ShapeType.roundRect,{x:.4,y:4.78,w:9.2,h:.58,fill:{color:LIGHT},rectRadius:.06,line:{color:LIGHT}});
  s.addText(msg,{x:.5,y:4.8,w:9,h:.54,fontFace:'Calibri',fontSize:10.5,color:INK,valign:'middle',margin:0,isTextBox:true});}
 s.addText('Mahisagar DTO • Data as of '+AS+' • '+sl,{x:6.5,y:5.4,w:3.1,h:.18,fontSize:7.5,color:MUT,align:'right',margin:0,isTextBox:true});return s;}
function tbl(s,head,rows,o){const hc=head.map(h=>({text:h,options:{bold:true,color:'FFFFFF',fill:{color:NAVY},align:'center',valign:'middle'}}));
 const body=rows.map((r,i)=>r.map((c,j)=>{const isO=(typeof c==='object'&&c!==null);const t=isO?c.t:c;
  return {text:String(t),options:{align:j==0?'left':'center',valign:'middle',bold:(j==0)||(i==rows.length-1&&o.totalRow),fill:{color:isO&&c.f?c.f:(i==rows.length-1&&o.totalRow?'DCE7EA':(i%2?'F7FAFB':'FFFFFF'))},color:INK}}}));
 s.addTable([hc,...body],{x:o.x,y:o.y,w:o.w,colW:o.colW,fontFace:'Calibri',fontSize:o.fs||9,rowH:o.rhs||o.rh||.27,border:{type:'solid',pt:.5,color:'D3DDE0'},margin:[0,.04,0,.04]});}
const card=(s,x,y,w,h,big,lab,col)=>{s.addShape(p.ShapeType.roundRect,{x,y,w,h,fill:{color:'FFFFFF'},rectRadius:.08,line:{color:'D3DDE0',width:1},shadow:{type:'outer',color:'000000',opacity:.1,blur:4,offset:2,angle:90}});
 s.addText(big,{x:x+.05,y:y+.08,w:w-.1,h:h*.5,fontFace:'Cambria',fontSize:h>1?26:22,bold:true,color:col||NAVY,align:'center',margin:0,isTextBox:true});
 s.addText(lab,{x:x+.08,y:y+h*.55,w:w-.16,h:h*.42,fontFace:'Calibri',fontSize:9,color:MUT,align:'center',valign:'top',margin:0,isTextBox:true});};
const chartOpt=(t,extra)=>Object.assign({showTitle:true,title:t,titleFontSize:11,titleColor:NAVY,catAxisLabelColor:MUT,valAxisLabelColor:MUT,catAxisLabelFontSize:9,valAxisLabelFontSize:9,valGridLine:{color:'E3EAEC',size:.5},catGridLine:{style:'none'},dataLabelFontSize:8,dataLabelColor:INK},extra);

const S9=a=>a.slice(0,9).reduce((x,y)=>x+y,0);
const TC=t=>t[0]+t.slice(1).toLowerCase();
const P1=R.pres['Total (District)'],N1=R.naat['Total (District)'],M1=R.micro['Total (District)'];
const NT=R.notif,DG=NT.diag,EN=NT.enrol,ND=S9(DG.District.m26),NE=S9(EN.District.m26),NTG=R.ntgt_month;
const NTOT=Object.values(NTG).reduce((a,b)=>a+b,0)*9;
const OA=R.out.all,OTOT=Object.values(OA).reduce((a,b)=>a+b,0);
const IPD=R.ip2,CPD=R.cp2,IPT=IPD.tot.District,CPT=CPD.tot.District;
const XD=R.xray.xray.District.slice(0,9).reduce((a,b)=>[a[0]+b[0],a[1]+b[1]],[0,0]);
const DM=R.dtc2.data,DFU=R.dtc2.fus;
const dsum=(m,t)=>DM[m][t].reduce((a,b)=>[a[0]+b[0],a[1]+b[1]],[0,0]);
const TPV=R.tpt['Mahisagar Total'];

// 1 TITLE
{const s=p.addSlide();s.background={color:NAVY};sl++;
 s.addText('MAHISAGAR DISTRICT • NTEP',{x:.6,y:.9,w:8,h:.3,fontSize:12,bold:true,color:'8FD3DC',charSpacing:3,margin:0,isTextBox:true});
 s.addText('District TB Programme\nPerformance Review 2026',{x:.6,y:1.3,w:8.5,h:1.5,fontFace:'Cambria',fontSize:38,bold:true,color:'FFFFFF',margin:0,isTextBox:true});
 s.addText('Presumptive • NAAT & Microscopy • X-ray • Notification (public / private / both bases) • Treatment initiation • Outcomes • IP/CP & PTFU • DTC (Jul · Aug · Sep by taluka) • Co-morbidity • TPT • NPY • ABHA',{x:.6,y:3.0,w:8.3,h:.8,fontSize:12.5,color:'CFE3E7',margin:0,isTextBox:true});
 s.addText('Data period: January – September 2026 (registers extracted '+AS+')  |  Prepared for District TB Programme Review, with expert commentary',{x:.6,y:4.5,w:8.8,h:.5,fontSize:11,color:'8FD3DC',margin:0,isTextBox:true});}

// 2 SCORECARD
{const s=base('Executive summary','District scorecard, Jan–Sep 2026','Bottom line: treatment outcomes are strong, but case-finding leaks at the testing step, the whole notification gap is in the private sector, and Lunawada is the weak link on most service indicators.');
 const pa=S9(P1.a),na=S9(N1.a);
 const k=[[pc(pa,P1.t*9,0),'Presumptive tested vs 9-month target ('+n(pa)+')',GRN],[pc(na,N1.t*9,0),'NAAT vs 9-month target; Sep only '+pc(N1.a[8],N1.t,0),RED],
 [pc(XD[0],XD[1],0),'X-ray offered ('+n(XD[0])+' of '+n(XD[1])+' enrolled)',AMB],[pc(ND,NTOT,0),'Notified vs pro-rata target ('+n(ND)+'); '+pc(ND-DG.District.ytd25,DG.District.ytd25,0)+' vs Jan–Sep 2025',RED],
 [pc(OA.S,OTOT,1),'Treatment success, 2025 cohort (9-month rule, n='+n(OTOT)+')',GRN],[pc(R.gap.late,R.gap.n,0),'Treatment started >7 days after diagnosis; '+R.ni.old+' never started',AMB],
 [pc(IPT[1],IPT[0],0),'IP follow-up done vs due',AMB],[pc(CPT[1],CPT[0],0),'CP follow-up done vs due',RED],
 [pc(dsum('Sep','District')[0],dsum('Sep','District')[1],0),'DTC done, Sep-26 (Aug '+pc(dsum('Aug','District')[0],dsum('Aug','District')[1],0)+', Jul '+pc(dsum('Jul','District')[0],dsum('Jul','District')[1],0)+')',AMB],
 [pc(TPV[3],TPV[2],0),'Contacts given TPT; '+pc(TPV[1],TPV[0],0)+' of pulmonary patients have zero contacts listed',AMB]];
 k.forEach((c,i)=>card(s,.4+(i%5)*1.86,.98+Math.floor(i/5)*1.8,1.72,1.6,c[0],c[1],c[2]));}

// 3 PRESUMPTIVE
{const sp=pc(P1.a[8],P1.t,0);const s=base('1 · Case-finding','Presumptive TB tested vs monthly target (no. and %)','Target is MONTHLY (district 3,032). Jan–Sep reaches '+pc(S9(P1.a),P1.t*9,0)+'; Santrampur ('+pc(S9(R.pres.Santrampur.a),R.pres.Santrampur.t*9,0)+') and Lunawada ('+pc(S9(R.pres.Lunawada.a),R.pres.Lunawada.t*9,0)+') hold the largest targets and trail most. Sep shows '+sp+' only because microscopy results jumped to '+n(M1.a[8])+' (next slide).');
 const cA=(a,t)=>{const x=100*a/t;return {t:n(a)+'\n'+x.toFixed(0)+'%',f:rag(x,90,70)}};
 const rows=[...T,'Total (District)'].map(t=>{const r=R.pres[t];return [t==='Total (District)'?'District':t,n(Math.round(r.t)),...r.a.slice(0,9).map(v=>cA(v,r.t)),cA(S9(r.a),r.t*9)]});
 tbl(s,['Taluka','Monthly target',...MS,'Jan–Sep total'],rows,{x:.4,y:.92,w:9.2,colW:[1.15,.7,...Array(9).fill(.66),1.41],totalRow:true,rhs:[.4,...Array(7).fill(.45)],fs:8.5});
 s.addText('Each cell: patients with NAAT or microscopy result, then % of monthly target (Jan–Sep total: % of 9-month target). Green ≥90% · Amber 70–89% · Red <70%.',{x:.4,y:4.5,w:9.2,h:.22,fontSize:8,color:MUT,margin:0,isTextBox:true});}

// 4 NAAT vs MICRO
{const s=base('1 · Case-finding','NAAT is falling while microscopy surges in September','NAAT reached only '+pc(N1.a[8],N1.t,0)+' of the Sep target while microscopy rose from '+M1.a[7]+' (Aug) to '+n(M1.a[8])+'. If NAAT cartridges/machines are short, patients are being diverted to the less sensitive test – verify stock and uptime now.');
 s.addChart(p.charts.BAR,[{name:'NAAT',labels:MS,values:N1.a.slice(0,9)},{name:'Microscopy',labels:MS,values:M1.a.slice(0,9)}],chartOpt('District – patients with result in month',{x:.4,y:.95,w:4.7,h:3.7,chartColors:[TEAL,AMB],showLegend:true,legendPos:'b',legendFontSize:9,barGapWidthPct:60}));
 const rows=T.map(t=>{const a=R.naat[t],m=R.micro[t],na=S9(a.a),v=100*na/(a.t*9),sv=100*a.a[8]/a.t;return [t,n(na),{t:v.toFixed(0)+'%',f:rag(v,90,70)},{t:sv.toFixed(0)+'%',f:rag(sv,70,50)},n(m.a[8])]});
 rows.push(['District',n(S9(N1.a)),pc(S9(N1.a),N1.t*9,0),pc(N1.a[8],N1.t,0),n(M1.a[8])]);
 tbl(s,['Taluka','NAAT Jan–Sep','% 9-mo target','NAAT Sep %','Micro Sep'],rows,{x:5.3,y:1.1,w:4.3,colW:[1,.85,.85,.8,.8],totalRow:true,rh:.38,fs:10});}

// 5 XRAY
{const X=R.xray.xray;const lo=T.map(t=>[t,X[t].slice(0,9).reduce((a,b)=>[a[0]+b[0],a[1]+b[1]],[0,0])]).map(([t,v])=>[t,100*v[0]/v[1]]).sort((a,b)=>a[1]-b[1]);
 const s=base('1 · Case-finding','X-ray offered to enrolled presumptives (number and %)','Every enrolled presumptive must get an X-ray. District rate is '+pc(XD[0],XD[1],0)+'; lowest are '+lo[0][0]+' ('+lo[0][1].toFixed(0)+'%), '+lo[1][0]+' ('+lo[1][1].toFixed(0)+'%) and '+lo[2][0]+' ('+lo[2][1].toFixed(0)+'%). September fell to '+pc(X.District[8][0],X.District[8][1],0)+' as enrolments rose to '+n(X.District[8][1])+'. Jan: 10,760 enrolled in a mass drive.');
 const hd=['Taluka',...MS,'Total'];
 const rows=[...T,'District'].map(t=>{const v=X[t].slice(0,9);const tot=v.reduce((a,b)=>[a[0]+b[0],a[1]+b[1]],[0,0]);
  return [t,...v.map(c=>{const q=c[1]?100*c[0]/c[1]:0;return {t:c[1]?n(c[0])+'/'+n(c[1])+'\n'+q.toFixed(0)+'%':'–',f:rag(q,80,65)}}),{t:n(tot[0])+'/'+n(tot[1])+'\n'+(100*tot[0]/tot[1]).toFixed(1)+'%',f:rag(100*tot[0]/tot[1],80,65)}]});
 tbl(s,hd,rows,{x:.4,y:1.0,w:9.2,colW:[1.05,...Array(9).fill(.78),1.13],totalRow:true,rhs:[.3,...Array(7).fill(.46)],fs:8});
 s.addText('Each cell: X-ray offered / presumptives enrolled that month, then %. Green ≥80% · Amber 65–79% · Red <65%. Enrolled Jan–Sep: '+n(XD[1])+'.',{x:.4,y:4.47,w:9,h:.25,fontSize:8.5,color:MUT,margin:0,isTextBox:true});}

// 6 CASCADE
{const X=R.xray;const s=base('1 · Case-finding','Where presumptives leak: the testing step and data quality','Only '+pc(X.tested,X.enrolled,0)+' of enrolled presumptives have a test result, and '+pc(X.asym_tested,X.tested,0)+' of those tested are labelled “Asymptomatic” – symptom screening is either not recorded correctly or testing is X-ray/risk-led without the field being updated.');
 card(s,.4,1.0,1.72,1.45,n(X.enrolled),'Presumptives enrolled (all 2026)',NAVY);card(s,2.25,1.0,1.72,1.45,pc(X.offered,X.enrolled,0),'X-ray offered',AMB);card(s,4.1,1.0,1.72,1.45,pc(X.result_recorded,X.offered,1),'X-ray result recorded ('+n(X.result_recorded)+' of '+n(X.offered)+' offered)',RED);
 card(s,5.95,1.0,1.72,1.45,pc(X.abn_tested,X.abn,0),'Abnormal X-ray tested by NAAT/microscopy ('+X.abn_tested+'/'+X.abn+')',GRN);card(s,7.8,1.0,1.8,1.45,pc(X.tested,X.enrolled,0),'Enrolled with NAAT or microscopy result ('+n(X.tested)+')',RED);
 s.addText([{text:'What this means',options:{bold:true,color:NAVY,breakLine:true}},
  {text:'X-ray is offered to about 8 in 10, but the result is almost never entered, so abnormal films cannot reliably trigger NAAT/microscopy – the '+pc(X.abn_tested,X.abn,0)+' linkage is only on the '+X.abn+' results recorded.',options:{bullet:true,breakLine:true}},
  {text:n(X.asym_tested)+' patients recorded as purely “Asymptomatic” have a test result – a wrongly-defined-asymptomatic list should be pulled by TU and corrected.',options:{bullet:true,breakLine:true}},
  {text:'The Jan mass drive ('+n(R.xray.xray.District[0][1])+' enrolled) inflates the base; overall only '+pc(X.tested,X.enrolled,0)+' of all enrolled presumptives have a test result.',options:{bullet:true}}],{x:.4,y:2.65,w:9.2,h:1.95,fontSize:11.5,color:INK,paraSpaceAfter:5,margin:0,isTextBox:true,valign:'top'});}

// 7 NOTIFICATION
{const kd=T.map(t=>[t,S9(DG[t].m26),NTG[t]*9]);
 const s=base('2 · Notification','Notification vs target – diagnosing and enrolment base','Notified '+n(ND)+' vs pro-rata target '+n(Math.round(NTOT))+' ('+pc(ND,NTOT,0)+'), '+Math.abs(Math.round(100*(ND-DG.District.ytd25)/DG.District.ytd25))+'% below Jan–Sep 2025. Lunawada is at '+pc(S9(DG.Lunawada.m26),NTG.Lunawada*9,0)+' of target; only Kadana exceeds it. Both bases are shown because either route earns the area credit.');
 let rows=T.map(t=>{const d=DG[t],e=EN[t],dn=S9(d.m26),en=S9(e.m26),pr=NTG[t]*9,yo=100*(dn-d.ytd25)/d.ytd25;
  return [t,Math.round(NTG[t]*12),Math.round(pr),dn,{t:(100*dn/pr).toFixed(0)+'%',f:rag(100*dn/pr,95,80)},en,{t:(100*en/pr).toFixed(0)+'%',f:rag(100*en/pr,95,80)},d.ytd25,{t:(yo>0?'+':'')+yo.toFixed(0)+'%',f:rag(yo+30,30,15)}]});
 const D=DG.District;rows.push(['District',Math.round(NTOT/9*12),Math.round(NTOT),ND,pc(ND,NTOT,0),NE,pc(NE,NTOT,0),D.ytd25,((ND-D.ytd25)*100/D.ytd25).toFixed(0)+'%']);
 tbl(s,['Taluka','Annual target','9-mo target','Diagnosed here','% target','Enrolled here','% target','Jan–Sep 2025','YoY (diag.)'],rows,{x:.4,y:1.0,w:9.2,colW:[1.2,.9,.9,1,.9,1,.9,1.2,1.2],totalRow:true,rh:.43,fs:10.5});}

// 8 PUBLIC vs PRIVATE vs TARGET
{const TG2=R.ntgt_pubpriv;const pubT=Object.values(TG2).reduce((a,b)=>a+b[0],0)*9,prvT=Object.values(TG2).reduce((a,b)=>a+b[1],0)*9;const pubA=DG.District.pub26,prvA=DG.District.prv26;
 const s=base('2 · Notification','Public vs private notification against TU-wise targets','Public sector delivers '+pc(pubA,pubT,0)+' of its 9-month target ('+n(pubA)+' of '+n(Math.round(pubT))+'); the gap is private: '+n(prvA)+' of '+n(Math.round(prvT))+' ('+pc(prvA,prvT,0)+'). Private notification fell from '+DG.District.prv25ytd+' (Jan–Sep 2025) to '+prvA+'.');
 const rows=T.map(t=>{const d=DG[t],pt=TG2[t][0]*9,vt=TG2[t][1]*9,pp=pt?100*d.pub26/pt:0,vp=vt?100*d.prv26/vt:0;
  return [t,Math.round(pt),d.pub26,{t:pt?pp.toFixed(0)+'%':'–',f:pt?rag(pp,95,80):undefined},d.pub25ytd,Math.round(vt),d.prv26,{t:vt?vp.toFixed(0)+'%':'no target',f:vt?rag(vp,80,50):undefined},d.prv25ytd]});
 rows.push(['District',Math.round(pubT),pubA,pc(pubA,pubT,0),DG.District.pub25ytd,Math.round(prvT),prvA,pc(prvA,prvT,0),DG.District.prv25ytd]);
 tbl(s,['Taluka','Public target (9 mo)','Public actual','% target','Public Jan–Sep 25','Private target (9 mo)','Private actual','% target','Private Jan–Sep 25'],rows,{x:.4,y:1.0,w:9.2,colW:[1.2,1,.9,.9,1,1.1,.9,1.1,1.1],totalRow:true,rh:.43,fs:10});}

// 9 FLOW
{const o=NT.out_of_diag,inn=NT.into_enrol;const net=t=>S9(EN[t].m26)-S9(DG[t].m26);
 const s=base('2 · Notification','Diagnosing vs enrolment facility: where patients move','Lunawada diagnoses '+o.Lunawada+' patients who are enrolled in other talukas (net '+net('Lunawada')+'); Khanpur (+'+net('Khanpur')+') and Santrampur (+'+net('Santrampur')+') gain on enrolment. '+NT.prv2pub+' privately diagnosed patients are enrolled in public units – the private sector feeds care but is not credited with notification.');
 s.addChart(p.charts.BAR,[{name:'Diagnosed at facility in taluka',labels:T,values:T.map(t=>S9(DG[t].m26))},{name:'Enrolled at facility in taluka',labels:T,values:T.map(t=>S9(EN[t].m26))}],chartOpt('Notified Jan–Sep 2026',{x:.4,y:.95,w:5.2,h:3.7,chartColors:[NAVY,AMB],showLegend:true,legendPos:'b',legendFontSize:9,showValue:true,dataLabelPosition:'outEnd',barGapWidthPct:50}));
 s.addText([{text:'Reading the two bases',options:{bold:true,color:NAVY,breakLine:true}},
 {text:'The presumptive register uses the enrolling facility: more enrolment → more diagnoses, either at your diagnosing facility or via your enrolling facility.',options:{bullet:true,breakLine:true}},
 {text:'Patients leaving diagnosing taluka: Lunawada '+o.Lunawada+', Kadana '+o.Kadana+'. Patients arriving: Santrampur '+inn.Santrampur+', Khanpur '+inn.Khanpur+'.',options:{bullet:true,breakLine:true}},
 {text:'Action: each TU tracks both numbers and refers patients diagnosed elsewhere to its own unit for enrolment.',options:{bullet:true}}],{x:5.8,y:1.0,w:3.8,h:3.6,fontSize:10.5,color:INK,paraSpaceAfter:6,margin:0,isTextBox:true,valign:'top'});}

// 10 TREATMENT INITIATION
{const nt=R.ni.tbu_old;const s=base('3 · Treatment initiation','Time from diagnosis to treatment start (target: within 7 days)','Median delay is only '+R.gap.med+' days, but '+pc(R.gap.late,R.gap.n,0)+' wait more than a week, and '+R.ni.old+' notified patients diagnosed before '+'30-Sep still have no treatment start date – Lunawada holds '+(nt.Lunawada||0)+' of them.');
 card(s,.4,1.0,2.1,1.35,pc(R.gap.late,R.gap.n,0),'initiated >7 days after diagnosis ('+R.gap.late+' of '+n(R.gap.n)+')',AMB);card(s,.4,2.5,2.1,1.35,String(R.ni.old),'no initiation date, outcome blank, diagnosed >7 days before register date',RED);
 const rows=T.map(t=>{const g=R.gap.tbu[t],v=100*g[0]/g[1];return [t,g[1],g[0],{t:v.toFixed(0)+'%',f:rag(100-v,85,75)},nt[t]||0]});
 rows.push(['District',R.gap.n,R.gap.late,pc(R.gap.late,R.gap.n,0),R.ni.old]);
 tbl(s,['Taluka (current TBU)','Started (Jan–Sep dx)','>7 days','% delayed','Not initiated >7d'],rows,{x:2.8,y:1.0,w:6.8,colW:[1.8,1.3,1,1.2,1.5],totalRow:true,rh:.42,fs:10.5});}

// 11 OUTCOME
{const dr=T.map(t=>{const o=R.out.tbu[t],N=Object.values(o).reduce((a,b)=>a+b,0);return [t,100*(o.D||0)/N]}).sort((a,b)=>b[1]-a[1]);
 const s=base('4 · Treatment outcome','Outcome of 2025 cohort – number / denominator and %','Only Jan–Dec 2025 notifications are counted (9-month rule). Success '+pc(OA.S,OTOT,1)+' clears the 90% benchmark; deaths '+pc(OA.D,OTOT,1)+' but highest in '+dr[0][0]+' ('+dr[0][1].toFixed(1)+'%) and '+dr[1][0]+' ('+dr[1][1].toFixed(1)+'%) – audit deaths there. DS-TB, 2HRZE/4HRE, pulmonary, by current facility.');
 [['S','Success',GRN],['D','Died',RED],['R','Regimen changed',AMB],['O','Other (LFU/failure/refused)',AMB],['P','Pending',NAVY]].forEach((c,i)=>card(s,.4+i*1.86,.95,1.72,1.05,pc(OA[c[0]],OTOT,1),c[1]+' ('+n(OA[c[0]])+'/'+n(OTOT)+')',c[2]));
 const oc=(a,N,hi,lo,inv)=>{const x=100*a/N;return {t:n(a)+'/'+n(N)+'\n'+x.toFixed(1)+'%',f:rag(inv?100-x:x,hi,lo)}};
 const rows=T.map(t=>{const o=R.out.tbu[t],N=Object.values(o).reduce((a,b)=>a+b,0);return [t,n(N),oc(o.S||0,N,92,90),oc(o.D||0,N,96,94,1),oc(o.R||0,N,98,96,1),oc(o.O||0,N,98,96,1),oc(o.P||0,N,99,97,1)]});
 rows.push(['District',n(OTOT),oc(OA.S||0,OTOT,92,90),oc(OA.D||0,OTOT,96,94,1),oc(OA.R||0,OTOT,98,96,1),oc(OA.O||0,OTOT,98,96,1),oc(OA.P||0,OTOT,99,97,1)]);
 tbl(s,['Taluka','Notified 2025 (denominator)','Success (n/N, %)','Died (n/N, %)','Regimen changed (n/N, %)','Other (n/N, %)','Pending (n/N, %)'],rows,{x:.4,y:2.1,w:9.2,colW:[1.3,1.3,1.3,1.3,1.4,1.3,1.3],totalRow:true,rhs:[.4,...Array(7).fill(.32)],fs:9});}

// 12 OUTCOME TREND
{const ML=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];const Ns=R.out.mon.map(o=>Object.values(o).reduce((a,b)=>a+b,0));
 const sv=R.out.mon.map((o,i)=>100*(o.S||0)/Ns[i]),dv=R.out.mon.map((o,i)=>100*(o.D||0)/Ns[i]);const lo=sv.indexOf(Math.min(...sv)),hi=dv.indexOf(Math.max(...dv));
 const e=R.out.early;const pend=e.map(o=>o.P||0);
 const s=base('4 · Treatment outcome','Outcome trend by month of notification and early-warning cohorts','Monthly success ranges '+Math.min(...sv).toFixed(0)+'–'+Math.max(...sv).toFixed(0)+'% (lowest '+ML[lo]+'-25); deaths peak at '+Math.max(...dv).toFixed(1)+'% in '+ML[hi]+'-25. Feb-26 already shows '+pend[1]+' and Mar-26 '+pend[2]+' outcomes pending – close them in Nikshay as 9-month windows mature.');
 s.addChart(p.charts.LINE,[{name:'Success %',labels:ML,values:sv.map(v=>+v.toFixed(1))},{name:'Died %',labels:ML,values:dv.map(v=>+v.toFixed(1))}],chartOpt(R.out.cohort_year+' notification cohort: success vs death %',{x:.4,y:.95,w:5.6,h:3.7,chartColors:[GRN,RED],showLegend:true,legendPos:'b',legendFontSize:9,lineSize:2.5,lineDataSymbolSize:6,valAxisMinVal:0,valAxisMaxVal:100}));
 const rows=e.map((o,i)=>{const N=Object.values(o).reduce((a,b)=>a+b,0);return [['Jan-26','Feb-26','Mar-26'][i],N,pc(o.S||0,N,0),o.P||0,{t:pc(o.P||0,N,0),f:rag(100-100*(o.P||0)/N,95,80)}]});
 tbl(s,['2026 cohort','Notified','Success','Pending','Pending %'],rows,{x:6.15,y:1.1,w:3.45,colW:[.85,.65,.7,.65,.6],rh:.4,fs:10});
 s.addText('Not yet mature (<9 months): reported for early action only, not as final outcome.',{x:6.15,y:2.85,w:3.45,h:.7,fontSize:9.5,italic:true,color:MUT,margin:0,isTextBox:true});}
// 13 IP
{const U=['Balasinor','Khanpur','Lunawada','Santrampur','Virpur','District'];const ml=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug*'];
 const cell=(v)=>{if(!v[0])return '–';const x=100*v[1]/v[0];return {t:v[1]+'/'+v[0]+'\n'+x.toFixed(0)+'%',f:rag(x,85,70)}};
 const low=U.slice(0,5).map(u=>[u,100*IPD.tot[u][1]/IPD.tot[u][0]]).sort((a,b)=>a[1]-b[1]);const jul=IPD.units.District[6];
 const s=base('5 · Treatment follow-up','IP follow-up – done / due and %, by month of diagnosis','IP follow-up is '+pc(IPT[1],IPT[0],0)+' overall; '+low[0][0]+' is '+low[0][1].toFixed(0)+'% with '+(IPD.tot.Lunawada[0]-IPD.tot.Lunawada[1])+' patients overdue. Jul cohort is '+pc(jul[1],jul[0],0)+' district-wide (some still within window). '+IPD.all_due_not_done+' patients are on the due-not-done line list.');
 const rows=U.map(u=>[u,...IPD.units[u].slice(0,8).map(cell),{t:IPD.tot[u][1]+'/'+IPD.tot[u][0]+'\n'+pc(IPD.tot[u][1],IPD.tot[u][0],1),f:rag(100*IPD.tot[u][1]/IPD.tot[u][0],85,70)}]);
 tbl(s,['TB Unit',...ml,'Total'],rows,{x:.4,y:1.0,w:9.2,colW:[1.6,...Array(8).fill(.8),1.2],totalRow:true,rhs:[.32,...Array(6).fill(.46)],fs:9.5});
 s.addText('Month of diagnosis. *Aug cohort mostly not yet due (>55 days). Denominator: pulmonary DS-TB due as on '+AS+'. Totals: '+IPT[0]+' due, '+IPT[1]+' done. Kadana (non-functional TBU) shown under Lunawada.',{x:.4,y:4.2,w:9.2,h:.5,fontSize:9,color:MUT,margin:0,isTextBox:true});}

// 14 CP
{const U=['Balasinor','Khanpur','Lunawada','Santrampur','Virpur','District'];const ml=['Sep-25','Oct-25','Nov-25','Dec-25','Jan-26','Feb-26','Mar-26','Apr-26*'];
 const cell=(v)=>{if(!v[0])return '–';const x=100*v[1]/v[0];return {t:v[1]+'/'+v[0]+'\n'+x.toFixed(0)+'%',f:rag(x,85,70)}};
 const s=base('5 · Treatment follow-up','CP follow-up – done / due and %, by month of diagnosis','CP follow-up is '+pc(CPT[1],CPT[0],0)+' overall and weakest in Lunawada ('+pc(CPD.tot.Lunawada[1],CPD.tot.Lunawada[0],0)+', '+(CPD.tot.Lunawada[0]-CPD.tot.Lunawada[1])+' overdue; Sep-25 cohort only '+pc(CPD.units.Lunawada[0][1],CPD.units.Lunawada[0][0],0)+'). Cohorts older than 6 months should already be ≥85%.');
 const rows=U.map(u=>[u,...CPD.units[u].slice(0,8).map(cell),{t:CPD.tot[u][1]+'/'+CPD.tot[u][0]+'\n'+pc(CPD.tot[u][1],CPD.tot[u][0],1),f:rag(100*CPD.tot[u][1]/CPD.tot[u][0],85,70)}]);
 tbl(s,['TB Unit',...ml,'Total'],rows,{x:.4,y:1.0,w:9.2,colW:[1.5,...Array(8).fill(.8),1.3],totalRow:true,rhs:[.32,...Array(6).fill(.46)],fs:9.5});
 s.addText('Month of diagnosis. *Apr-26 cohort only partly due. Denominator: pulmonary DS-TB due for CP follow-up as on '+AS+'. Totals: '+CPT[0]+' due, '+CPT[1]+' done; '+CPD.all_due_not_done+' on the due-not-done list.',{x:.4,y:4.2,w:9.2,h:.5,fontSize:9,color:MUT,margin:0,isTextBox:true});}

// 15 PTFU
{const PT=R.ptfu;const mm=PT.District.map(x=>x[0]);const dist=PT.District;const sp=dist[3];const lun=PT.Lunawada[3];
 const s=base('5 · Treatment follow-up','Post-treatment follow-up (6/12/18/24 months) – done / due and %','District PTFU is '+sp[3]+'% in Sep-26 ('+sp[1]+' due, '+sp[2]+' done, '+(sp[1]-sp[2])+' pending). Lunawada is '+lun[3]+'% and holds '+(lun[1]-lun[2])+' of the pending; Khanpur ('+PT.Khanpur[3][3]+'%) and Santrampur ('+PT.Santrampur[3][3]+'%) are near-complete. Jun-26 is low ('+dist[0][3]+'%) across the district.');
 s.addChart(p.charts.BAR,[{name:'PTFU %',labels:mm.map(x=>x.split('-')[0]),values:dist.map(x=>x[3])}],chartOpt('District PTFU achievement %',{x:.4,y:.95,w:3.6,h:3.7,chartColors:[TEAL],showValue:true,dataLabelPosition:'outEnd',showLegend:false,valAxisMaxVal:100,valAxisMinVal:0}));
 const rows=[...T,'District'].map(k=>[k,...PT[k].map(x=>({t:n(x[2])+'/'+n(x[1])+'\n'+x[3]+'%',f:rag(x[3],85,70)}))]);
 tbl(s,['Taluka',...mm],rows,{x:4.2,y:1.0,w:5.4,colW:[1.2,1.05,1.05,1.05,1.05],totalRow:true,rhs:[.32,...Array(7).fill(.46)],fs:9.5});
 s.addText('Cell: PTFU done / due, then %. Due = outcome 6/12/18/24 months earlier.',{x:4.2,y:4.5,w:5.4,h:.22,fontSize:8,color:MUT,margin:0,isTextBox:true});}

// 16 DTC overview
{const mo=['Jul','Aug','Sep'];const T3=t=>mo.map(m=>dsum(m,t));
 const s=base('6 · Differentiated TB care','DTC completion: Jul, Aug, Sep – done / due and %','DTC done: Jul '+pc(...dsum('Jul','District'),0)+', Aug '+pc(...dsum('Aug','District'),0)+', Sep '+pc(...dsum('Sep','District'),0)+'. Baseline completion fell from '+pc(...DM.Jul.District[0],0)+' (Jul) to '+pc(...DM.Sep.District[0],0)+' (Sep). Khanpur leads in Aug–Sep; Lunawada trails at '+pc(...dsum('Sep','Lunawada'),0)+' in Sep. Overall DTC rose from '+Math.round(100*DM.Jan.District.reduce((a,v)=>a+v[0],0)/DM.Jan.District.reduce((a,v)=>a+v[1],0))+'% (Jan) to '+Math.round(100*DM.Jun.District.reduce((a,v)=>a+v[0],0)/DM.Jun.District.reduce((a,v)=>a+v[1],0))+'% (Jun). Taluka detail on the next three slides.');
 s.addChart(p.charts.BAR,[0,1,2].map(i=>({name:mo[i]+'-26',labels:['Base','1st','2nd','3rd','4th','5th','6th'],values:DM[mo[i]].District.map(v=>Math.round(100*v[0]/v[1]))})),chartOpt('District DTC done % by follow-up type',{x:.3,y:.95,w:4.8,h:3.7,chartColors:['9FC5CC',TEAL,NAVY],showLegend:true,legendPos:'b',legendFontSize:9,barGapWidthPct:40,valAxisMaxVal:100,valAxisMinVal:0}));
 const rows=[...T,'District'].map(t=>[t,...T3(t).map(v=>{const x=100*v[0]/v[1];return {t:n(v[0])+'/'+n(v[1])+'\n'+x.toFixed(0)+'%',f:rag(x,75,60)}})]);
 tbl(s,['Taluka','Jul-26','Aug-26','Sep-26'],rows,{x:5.2,y:1.0,w:4.4,colW:[1.1,1.1,1.1,1.1],totalRow:true,rhs:[.32,...Array(7).fill(.46)],fs:9.5});}

// 17-19 DTC by month
const FUL=['Baseline','1st FU','2nd FU','3rd FU','4th FU','5th FU','6th FU'];
function dtcSlide(m,name){
 const tt=[...T,'District'].map(t=>[t,dsum(m,t)]);const dist=dsum(m,'District');
 const tal=tt.slice(0,6).map(([t,v])=>[t,100*v[0]/v[1]]).sort((a,b)=>b[1]-a[1]);const fu=DM[m].District.map((v,i)=>[FUL[i],100*v[0]/v[1]]).sort((a,b)=>a[1]-b[1]);
 const s=base('6 · Differentiated TB care','DTC by type, taluka-wise – '+name+' 2026 (done / due and %)',pc(...dist,0)+' of '+n(dist[1])+' due DTC assessments were done in '+name+'. Best: '+tal[0][0]+' ('+tal[0][1].toFixed(0)+'%); lowest: '+tal[5][0]+' ('+tal[5][1].toFixed(0)+'%). Weakest type district-wide: '+fu[0][0]+' ('+fu[0][1].toFixed(0)+'%). Due = patients whose diagnosis date falls in the window for that DTC type.');
 const rows=[...T,'District'].map(t=>{const tot=dsum(m,t);return [t,...DM[m][t].map(v=>{const x=v[1]?100*v[0]/v[1]:0;return {t:v[0]+'/'+v[1]+'\n'+x.toFixed(0)+'%',f:rag(x,80,60)}}),{t:tot[0]+'/'+tot[1]+'\n'+(100*tot[0]/tot[1]).toFixed(0)+'%',f:rag(100*tot[0]/tot[1],80,60)}]});
 tbl(s,['Taluka',...FUL,'All types'],rows,{x:.4,y:.9,w:9.2,colW:[1.2,...Array(7).fill(.95),1.35],totalRow:true,rhs:[.3,...Array(7).fill(.48)],fs:9.5});
 s.addText('Green ≥80% · Amber 60–79% · Red <60%.',{x:.4,y:4.55,w:5,h:.2,fontSize:8,color:MUT,margin:0,isTextBox:true});}
dtcSlide('Jul','July');dtcSlide('Aug','August');dtcSlide('Sep','September');

// 20 CO-MORBIDITY
{const C=R.co;const s=base('7 · Co-morbidity','HIV, diabetes and substance-use care among 2026 patients','HIV status is known for '+pc(C.n-C.hiv_pend,C.n,0)+' and diabetes status for '+pc(C.n-C.dm_pend,C.n,0)+', but only '+C.dm_init+' of '+C.dm+' diabetic patients ('+pc(C.dm_init,C.dm,0)+') have an anti-diabetic start recorded and '+C.tob_link+' of '+C.tob+' tobacco users ('+pc(C.tob_link,C.tob,0)+') are linked to cessation. Santrampur ('+C.dm_tbu.Santrampur[1]+' of '+C.dm_tbu.Santrampur[0]+') and Virpur ('+C.dm_tbu.Virpur[1]+' of '+C.dm_tbu.Virpur[0]+') are lowest on diabetes care.');
 [[pc(C.n-C.hiv_pend,C.n,1),'HIV status known (pending '+C.hiv_pend+')',GRN],[pc(C.n-C.dm_pend,C.n,1),'Diabetes status known (pending '+C.dm_pend+')',AMB],[String(C.hiv_pos),'HIV reactive/positive',NAVY],[pc(C.dm_init,C.dm,0),'Diabetics with anti-diabetic start ('+C.dm_init+'/'+C.dm+')',RED],[pc(C.tob_link,C.tob,0),'Tobacco users linked to cessation ('+C.tob_link+'/'+C.tob+')',RED]]
  .forEach((c,i)=>card(s,.4+i*1.86,.95,1.72,1.3,c[0],c[1],c[2]));
 const rows=T.map(t=>{const c=C.tbu[t],d=C.dm_tbu[t];return [t,c[0],{t:c[1],f:rag(100-100*c[1]/c[0],98,96)},{t:c[2],f:rag(100-100*c[2]/c[0],96,94)},d[0]?d[1]+'/'+d[0]:'–']});rows.push(['District',C.n,C.hiv_pend,C.dm_pend,C.dm_init+'/'+C.dm]);
 tbl(s,['Taluka','Patients 2026','HIV status pending','Diabetes status pending','Anti-diabetic started / diabetic'],rows,{x:.9,y:2.45,w:8.2,colW:[1.7,1.3,1.6,1.8,1.8],totalRow:true,rh:.28,fs:10});}

// 21 TPT
{const Tp=R.tpt;const lun=Tp.Lunawada,pv=Tp['All - Private'];
 const s=base('8 · TB preventive treatment','TPT: contact tracing is the bottleneck, not drug delivery','TPT is given to '+pc(TPV[3],TPV[2],0)+' of listed contacts, but '+pc(TPV[1],TPV[0],0)+' of pulmonary patients have no contact listed – '+pc(lun[1],lun[0],0)+' in Lunawada (also lowest TPT given at '+pc(lun[3],lun[2],0)+') and '+pc(pv[1],pv[0],0)+' among privately diagnosed patients. Sep contacts are still being traced.');
 const rows=[...T].map(t=>{const v=Tp[t],z=100*v[1]/v[0],g=100*v[3]/v[2];return [t,v[0],{t:v[1]+' ('+z.toFixed(0)+'%)',f:rag(100-z,80,65)},v[2],{t:v[3]+' ('+g.toFixed(0)+'%)',f:rag(g,90,80)}]});
 rows.push(['District',TPV[0],TPV[1]+' ('+pc(TPV[1],TPV[0],0)+')',TPV[2],TPV[3]+' ('+pc(TPV[3],TPV[2],0)+')']);
 tbl(s,['Taluka','Pulmonary patients','Zero contact (%)','Contacts listed','Given TPT (%)'],rows,{x:.4,y:1.0,w:5.9,colW:[1.2,1.1,1.3,1.1,1.2],totalRow:true,rh:.42,fs:10});
 const mm=Tp.monthly.slice(0,8);s.addChart(p.charts.LINE,[{name:'Given TPT %',labels:MS.slice(0,8),values:mm.map(m=>+(100*m[3]/m[2]).toFixed(0))}],chartOpt('Monthly TPT given % (Jan–Aug)',{x:6.4,y:.95,w:3.2,h:3.7,chartColors:[TEAL],showLegend:false,lineSize:2.5,valAxisMinVal:60,valAxisMaxVal:100,showValue:true,dataLabelPosition:'t'}));}

// 22 NPY
{const all={};T.forEach(t=>Object.entries(R.ben[t]).forEach(([k,c])=>all[k]=(all[k]||0)+c));const N=Object.values(all).reduce((a,c)=>a+c,0);const L=R.ben.Lunawada,NL=Object.values(L).reduce((a,c)=>a+c,0);
 const rj=Object.entries(R.ben_reject_top);const bank=rj.filter(x=>x[0].startsWith('Bank Name And IFSC')).reduce((a,x)=>a+x[1],0);
 const s=base('9 · Nikshay Poshan Yojana','Beneficiary bank validation (number and %)','In this register of '+n(N)+' beneficiaries, '+pc(all.VALIDATED,N,0)+' are bank-validated; Lunawada has '+L.EMPTY+' records with no bank details ('+pc(L.EMPTY,NL,0)+'). Most rejections are bank/IFSC mismatches ('+bank+' of '+all.REJECTED+'). '+(R.ben_pay.paid>0?'₹'+(R.ben_pay.paid/1e7).toFixed(2)+' cr paid against ₹'+(R.ben_pay.elig/1e7).toFixed(2)+' cr eligible ('+pc(R.ben_pay.paid,R.ben_pay.elig,0)+').':'The extract shows ₹'+(R.ben_pay.elig/1e7).toFixed(2)+' cr eligible and ₹0 paid through Nikshay – it looks like a pending-payment list, so confirm before quoting payment coverage.')+'');
 const cn=(a,N,hi,lo,inv)=>{const x=100*a/N;return {t:n(a)+'\n'+x.toFixed(0)+'%',f:inv?rag(100-x,hi,lo):rag(x,hi,lo)}};
 const ip=b=>(b.SENT||0)+(b.ENTERED||0)+(b.NOT_VALIDATED||0);
 const rows=T.map(t=>{const b=R.ben[t],Nn=Object.values(b).reduce((a,c)=>a+c,0);return [t,n(Nn),cn(b.VALIDATED||0,Nn,85,75),cn(b.EMPTY||0,Nn,95,85,1),cn(b.REJECTED||0,Nn,95,90,1),cn(ip(b),Nn,98,95,1)]});
 rows.push(['District',n(N),cn(all.VALIDATED||0,N,85,75),cn(all.EMPTY||0,N,95,85,1),cn(all.REJECTED||0,N,95,90,1),cn(ip(all),N,98,95,1)]);
 tbl(s,['Taluka','Beneficiaries','Validated (n, %)','Empty – no bank details (n, %)','Rejected (n, %)','In process (n, %)'],rows,{x:.4,y:1.0,w:9.2,colW:[1.6,1.3,1.5,1.7,1.55,1.55],totalRow:true,rhs:[.4,...Array(7).fill(.44)],fs:10});}

// 23 ABHA + CONSENT
{const A=R.abha,C=R.consent;const at=Object.values(A).reduce((a,b)=>[a[0]+b[0],a[1]+b[1]],[0,0]),ct=Object.values(C).reduce((a,b)=>[a[0]+b[0],a[1]+b[1]],[0,0]);const lo=T.map(t=>[t,100*A[t][0]/A[t][1]]).sort((a,b)=>a[1]-b[1]);
 const s=base('10 · PM-TBMBA','ABHA ID seeding and consent for Ni-kshay Mitra food kit','ABHA ID is generated for '+pc(at[0],at[1],0)+' of the '+n(at[1])+' patients in the PM-TBMBA register; lowest are '+lo[0][0]+' ('+lo[0][1].toFixed(0)+'%), '+lo[1][0]+' ('+lo[1][1].toFixed(0)+'%) and '+lo[2][0]+' ('+lo[2][1].toFixed(0)+'%). Consent for the kit is '+pc(ct[0],ct[1],0)+' but only '+pc(C.Kadana[0],C.Kadana[1],0)+' in Kadana.');
 const c2=(a,N,hi,lo,inv)=>{const x=100*a/N;return {t:n(a)+'\n'+x.toFixed(0)+'%',f:inv?rag(100-x,hi,lo):rag(x,hi,lo)}};
 const rows=T.map(t=>[t,n(A[t][1]),c2(A[t][0],A[t][1],70,55),c2(A[t][1]-A[t][0],A[t][1],45,30,1),c2(C[t][0],C[t][1],90,80),c2(C[t][1]-C[t][0],C[t][1],10,20,1)]);
 rows.push(['District',n(at[1]),c2(at[0],at[1],70,55),c2(at[1]-at[0],at[1],45,30,1),c2(ct[0],ct[1],90,80),c2(ct[1]-ct[0],ct[1],10,20,1)]);
 tbl(s,['Taluka','PM-TBMBA patients','ABHA generated (n, %)','ABHA not generated (n, %)','Kit consent given (n, %)','Consent not given (n, %)'],rows,{x:.4,y:1.0,w:9.2,colW:[1.6,1.3,1.6,1.6,1.55,1.55],totalRow:true,rhs:[.4,...Array(7).fill(.44)],fs:10});}

// 24 EXPERT
{const s=base('Expert assessment','Where Mahisagar stands – an honest read of the data','Priority order: (1) restore NAAT access and close the testing gap, (2) recover private-sector notification, (3) a dedicated Lunawada recovery plan.');
 const col=(x,h,c,items)=>{s.addShape(p.ShapeType.roundRect,{x,y:.95,w:2.95,h:3.7,fill:{color:'FFFFFF'},rectRadius:.08,line:{color:c,width:1.5}});
  s.addText(h,{x:x+.12,y:1.02,w:2.7,h:.3,fontFace:'Cambria',fontSize:13,bold:true,color:c,margin:0,isTextBox:true});
  s.addText(items.map((t,i)=>({text:t,options:{bullet:true,breakLine:i<items.length-1}})),{x:x+.12,y:1.38,w:2.72,h:3.2,fontSize:9.5,color:INK,paraSpaceAfter:5,margin:0,isTextBox:true,valign:'top'});};
 col(.4,'Doing well',GRN,['Outcomes: '+pc(OA.S,OTOT,1)+' success on the 2025 cohort; only '+OA.P+' pending','Median '+R.gap.med+' days to treatment start','Public sector at '+pc(DG.District.pub26,Object.values(R.ntgt_pubpriv).reduce((a,b)=>a+b[0],0)*9,0)+' of its notification target','HIV status known for '+pc(R.co.n-R.co.hiv_pend,R.co.n,0),'TPT given to '+pc(TPV[3],TPV[2],0)+' of listed contacts','Khanpur and Virpur: IP, CP and DTC at or near target; Khanpur PTFU '+R.ptfu.Khanpur[3][3]+'%']);
 col(3.55,'Needs attention',AMB,['Notification '+pc(ND,NTOT,0)+' of pro-rata and '+Math.abs(Math.round(100*(ND-DG.District.ytd25)/DG.District.ytd25))+'% below 2025','Private sector at '+pc(DG.District.prv26,Object.values(R.ntgt_pubpriv).reduce((a,b)=>a+b[1],0)*9,0)+' of its target','Only '+pc(R.xray.tested,R.xray.enrolled,0)+' of enrolled presumptives tested; '+pc(R.xray.asym_n,R.xray.enrolled,0)+' recorded asymptomatic','X-ray result entered for just '+pc(R.xray.result_recorded,R.xray.offered,1),pc(R.gap.late,R.gap.n,0)+' start treatment >7 days late; '+R.ni.old+' never started','Only '+R.co.dm_init+' of '+R.co.dm+' diabetics on treatment; '+R.co.tob_link+' of '+R.co.tob+' tobacco users linked']);
 col(6.7,'Critical – act now',RED,['Sep NAAT '+pc(N1.a[8],N1.t,0)+' of target while microscopy rose to '+n(M1.a[8])+' ('+Math.round(M1.a[8]/M1.a[7])+'× Aug) – likely supply/access shift','Lunawada is weakest on IP ('+pc(IPD.tot.Lunawada[1],IPD.tot.Lunawada[0],0)+'), CP ('+pc(CPD.tot.Lunawada[1],CPD.tot.Lunawada[0],0)+'), PTFU ('+R.ptfu.Lunawada[3][3]+'% Sep), DTC ('+pc(...dsum('Sep','Lunawada'),0)+' Sep), TPT zero-contact ('+pc(R.tpt.Lunawada[1],R.tpt.Lunawada[0],0)+') and NPY – a system problem, not isolated lapses','Deaths: Kadana 6.5%, Balasinor 5.4%','DTC fell to '+pc(...dsum('Sep','District'),0)+' in Sep']);}

// 25 ACTION PLAN
{const s=base('Action plan','Priority actions for Oct–Dec 2026',null);
 const A=[['1','Audit CBNAAT/Truenat cartridge stock and machine uptime; restore NAAT-first testing','DTO, STS, Lab technicians','7 days','NAAT ≥85% of monthly target'],
 ['2','Mandatory result entry for every X-ray; same-day NAAT for abnormal film; correct wrongly-asymptomatic entries','MO-TC, Radiographers','30 days','X-ray result ≥80%; ≥90% abnormal tested'],
 ['3','Private-sector drive (PPSA/PPM) in Lunawada, Santrampur and Balasinor to close the private target gap','DTO, PPSA Coordinator','45 days','Private ≥80% of target'],
 ['4','Dual-base review: TUs refer patients diagnosed elsewhere to their own unit for enrolment','MO-TC, Data manager','Monthly','Track diag. vs enrol base per TU'],
 ['5','Treat-within-7-days drive; trace the '+R.ni.old+' patients with no start date (Lunawada '+(R.ni.tbu_old.Lunawada||0)+', Santrampur '+(R.ni.tbu_old.Santrampur||0)+', Balasinor '+(R.ni.tbu_old.Balasinor||0)+')','STS, TB-HV','14 days','<10% delayed; 0 untreated >7d'],
 ['6','Lunawada recovery team: IP/CP/PTFU/DTC due-lists, TPT contact tracing, bank-detail camp ('+R.ben.Lunawada.EMPTY+' empty records)','District TB Officer, MO-TC Lunawada','60 days','IP/CP/PTFU ≥85%; zero-contact <15%'],
 ['7','Death audit in Kadana and Balasinor; close Feb–Mar-26 pending outcomes in Nikshay','MO-TC, STS','30 days','Pending <10%; death <3.5%'],
 ['8','Start anti-diabetic and cessation linkage; complete HIV/diabetes status for '+R.co.dm_pend+' pending; ABHA camps in Kadana, Virpur, Lunawada','Medical Officers, Counsellors','45 days','≥80% DM treated; ABHA ≥70%']];
 tbl(s,['#','Action','Responsible','Timeline','Target'],A,{x:.4,y:.95,w:9.2,colW:[.35,4.3,1.9,.8,1.85],rh:.47,fs:9});}

// 26 NOTES
{const s=base('Notes','Data notes and corrections to the previous version',null);
 s.addText([{text:'Corrected / changed in this update',options:{bold:true,color:NAVY,breakLine:true}},
 {text:'All indicators rebuilt from the registers extracted on '+AS+' (latest diagnosis '+RA+'); Sep-26 is now complete for notification and testing.',options:{bullet:true,breakLine:true}},
 {text:'TPT: contacts = persons attached to distinct pulmonary patients ('+n(TPV[2])+' contacts, '+pc(TPV[3],TPV[2],0)+' given). An earlier version of this deck overstated contacts.',options:{bullet:true,breakLine:true}},
 {text:'DTC: rule re-derived from the DTC register (figures can differ slightly from the earlier DTC workbook); Jul, Aug and Sep are shown by taluka. Done = DTC assessment of the matching type recorded; due = diagnosis date within the 28-day window for that type (baseline 0 days; 1st follow-up 56 days, then +28 days each). Overall DTC done by month: Jan '+Math.round(100*DM.Jan.District.reduce((a,v)=>a+v[0],0)/DM.Jan.District.reduce((a,v)=>a+v[1],0))+'%, Jun '+Math.round(100*DM.Jun.District.reduce((a,v)=>a+v[0],0)/DM.Jun.District.reduce((a,v)=>a+v[1],0))+'%.',options:{bullet:true,breakLine:true}},
 {text:'Notification targets now use your TU-wise public/private targets; presumptive/NAAT targets remain MONTHLY.',options:{bullet:true,breakLine:true}},
 {text:'Limits',options:{bold:true,color:NAVY,breakLine:true}},
 {text:'Kadana (non-functional) TBU is folded into Lunawada in IP/CP; no Private TB Unit row appears in this register. PTFU uses the PTFU cumulative file as uploaded (check its extract date); Jun-26 PTFU can differ from earlier workbooks because the register holds more outcome records.',options:{bullet:true,breakLine:true}},
 {text:'ABHA comes from the PM-TBMBA register ('+n(Object.values(R.abha).reduce((a,b)=>a+b[1],0))+' patients notified in 2026) and is not comparable with the old 31% figure. Beneficiary register: '+n(Object.values(R.ben).reduce((a,b)=>a+Object.values(b).reduce((x,y)=>x+y,0),0))+' records'+(R.ben_pay.paid>0?'.':' with no Nikshay payments, so NPY percentages are not comparable with earlier versions.'),options:{bullet:true}}],{x:.5,y:1.0,w:9,h:3.9,fontSize:11,color:INK,paraSpaceAfter:5,margin:0,isTextBox:true,valign:'top'});}

p.writeFile({fileName:OUTF}).then(()=>console.log('written',sl));

// Usage: NODE_PATH=$(npm root -g) node build_monthly_deck.js <monthly.json> <R.json from extract_data.py> <out.pptx>
const pptxgen=require('pptxgenjs');const M=require(require('path').resolve(process.argv[2]||'/tmp/monthly.json')),R=require(require('path').resolve(process.argv[3]||'/tmp/R.json'));const OUTF=process.argv[4]||'/mnt/user-data/outputs/Mahisagar_Monthly_Review.pptx';
const p=new pptxgen();p.layout='LAYOUT_16x9';p.title='Mahisagar District TB Programme – Monthly Review '+M.month;
const NAVY='0B3C49',TEAL='1B7F8C',LIGHT='EEF4F5',INK='1F2A30',MUT='5F6F76',GOOD='1E7B45',BAD='C0392B';
const T=['Balasinor','Kadana','Khanpur','Lunawada','Santrampur','Virpur'];
const n=x=>Number(x).toLocaleString('en-IN');
const MN=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];
const lab=per=>{const [y,m]=per.split('-');return MN[+m-1]+'-'+y.slice(2)};
const L=lab(M.periods.last),PV=lab(M.periods.prev),LY=lab(M.periods.lastyear);
const AS=new Date(R.asof).toLocaleDateString('en-GB',{day:'2-digit',month:'short',year:'numeric'}).replace(/ /g,'-');
let sl=0;
function base(sec,title,msg){const s=p.addSlide();s.background={color:'FFFFFF'};sl++;
 s.addText(sec.toUpperCase(),{x:.4,y:.14,w:7,h:.22,fontSize:9,bold:true,color:TEAL,margin:0,isTextBox:true});
 s.addText(title,{x:.4,y:.34,w:9.2,h:.5,fontFace:'Cambria',fontSize:21,bold:true,color:NAVY,margin:0,isTextBox:true});
 if(msg){s.addShape(p.ShapeType.roundRect,{x:.4,y:4.78,w:9.2,h:.58,fill:{color:LIGHT},rectRadius:.06,line:{color:LIGHT}});
  s.addText(msg,{x:.5,y:4.8,w:9,h:.54,fontFace:'Calibri',fontSize:10.5,color:INK,valign:'middle',margin:0,isTextBox:true});}
 s.addText('Mahisagar DTO • Monthly review '+L+' • Data as of '+AS+' • '+sl,{x:5.2,y:5.4,w:4.4,h:.18,fontSize:7.5,color:MUT,align:'right',margin:0,isTextBox:true});return s;}
function tbl(s,head,rows,o){const hc=head.map(h=>({text:h,options:{bold:true,color:'FFFFFF',fill:{color:NAVY},align:'center',valign:'middle'}}));
 const body=rows.map((r,i)=>r.map((c,j)=>{const isO=(typeof c==='object'&&c!==null);const t=isO?c.t:c;const last=i===rows.length-1&&o.totalRow;
  return {text:String(t),options:{align:j==0?'left':'center',valign:'middle',bold:(j==0)||last||(isO&&c.b),color:isO&&c.c?c.c:INK,fill:{color:last?'DCE7EA':(i%2?'F7FAFB':'FFFFFF')}}}}));
 s.addTable([hc,...body],{x:.4,y:o.y||.92,w:9.2,colW:o.colW,fontFace:'Calibri',fontSize:o.fs||9.5,rowH:[o.hh||.42,...Array(rows.length).fill(o.rh||.4)],border:{type:'solid',pt:.5,color:'D3DDE0'},margin:[0,.04,0,.04]});}
// helpers
const cell=(a,b)=>({t:a+'\n'+b});
const pct=(a,b,d=1)=>b?(100*a/b).toFixed(d)+'%':'–';
function chgPP(a,b,hb){if(a==null||b==null||isNaN(a)||isNaN(b))return {t:'n/a',c:MUT};const d=a-b;if(Math.abs(d)<.05)return {t:'= 0.0 pp',c:MUT};return {t:(d>0?'▲ +':'▼ ')+d.toFixed(1)+' pp',c:(d>0)===hb?GOOD:BAD,b:true}}
function chgN(a,b,hb){if(!b)return {t:'n/a',c:MUT};const d=100*(a-b)/b;if(Math.abs(d)<.5)return {t:'= 0%',c:MUT};return {t:(d>0?'▲ +':'▼ ')+d.toFixed(0)+'%',c:(d>0)===hb?GOOD:BAD,b:true}}
// rank markers: best two / lowest two among talukas with denominator >=10
function marks(vals,dens,hb){const ok=T.map((t,i)=>[t,vals[i],dens[i]]).filter(x=>x[1]!=null&&x[2]>=10).sort((a,b)=>hb?b[1]-a[1]:a[1]-b[1]);
 const best=ok.slice(0,2).map(x=>x[0]),worst=ok.slice(-2).map(x=>x[0]);return t=>best.includes(t)?{t:'▲ '+t,c:GOOD,b:true}:worst.includes(t)?{t:'▼ '+t,c:BAD,b:true}:t;}
const leg='▲ best two · ▼ lowest two talukas on the main % (denominator ≥10). Arrows: green = moved in the good direction, red = wrong direction.';
const P=M.presumptive,N=M.notification,TS=M.treatment_start,O=M.outcome_matured_cohort;
const lastName=t=>t;

// 1 TITLE
{const s=p.addSlide();s.background={color:NAVY};sl++;
 s.addText('MAHISAGAR DISTRICT • NTEP',{x:.6,y:.9,w:8,h:.3,fontSize:12,bold:true,color:'8FD3DC',charSpacing:3,margin:0,isTextBox:true});
 s.addText('Monthly Performance Review\n'+new Date(M.month+'-01').toLocaleDateString('en-GB',{month:'long',year:'numeric'}),{x:.6,y:1.3,w:8.5,h:1.5,fontFace:'Cambria',fontSize:38,bold:true,color:'FFFFFF',margin:0,isTextBox:true});
 s.addText('Every indicator: '+L+' vs '+PV+', vs monthly target, and vs '+LY+' – by taluka and for the district',{x:.6,y:3.0,w:8.3,h:.7,fontSize:13,color:'CFE3E7',margin:0,isTextBox:true});
 s.addText('Registers extracted '+AS+' (latest diagnosis '+new Date(R.reg_asof).toLocaleDateString('en-GB',{day:'2-digit',month:'short',year:'numeric'}).replace(/ /g,'-')+') – '+L+' is a complete month.',{x:.6,y:4.5,w:8.8,h:.5,fontSize:11,color:'8FD3DC',margin:0,isTextBox:true});}

// district shortcuts
const pd=P.last.District,pp=P.prev.District,nd=N.diag.last.District,np=N.diag.prev.District,nl=N.diag.lastyear.District;
const tl=TS.last.District,tp=TS.prev.District,tly=TS.lastyear.District,ol=O.last.District,op=O.prev.District;
const PT=R.ptfu.District,ptL=PT[3],ptP=PT[2];
const dS=(m,t)=>R.dtc2.data[m][t].reduce((a,b)=>[a[0]+b[0],a[1]+b[1]],[0,0]);const dL=dS('Sep','District'),dP=dS('Aug','District');
const nai=pd.naat/pd.tested*100,naiP=pp.naat/pp.tested*100;

// 2 SCORECARD
{const s=base('Scorecard','District scorecard – '+L,'Testing and notification slipped while treatment-start delay and deaths in the matured cohort improved. NAAT '+n(pd.naat)+' vs '+n(pp.naat)+' as microscopy rose '+Math.round(pd.micro/pp.micro)+'×; check cartridge/machine access first.');
 const row=(name,a,b,ch,tg,ly,lc)=>[name,a,b,ch,tg,ly,lc];
 const rows=[
 row('Presumptive enrolled (n)',n(pd.enrolled),n(pp.enrolled),chgN(pd.enrolled,pp.enrolled,true),'–','n/a',{t:'n/a',c:MUT}),
 row('Presumptive tested (by enrolment month)',n(pd.tested)+' ('+pct(pd.tested,pd.enrolled,0)+')',n(pp.tested)+' ('+pct(pp.tested,pp.enrolled,0)+')',chgN(pd.tested,pp.tested,true),pd.tested_pct_target.toFixed(0)+'%','n/a',{t:'n/a',c:MUT}),
 row('NAAT done (n, % of target)',n(pd.naat)+' ('+pd.naat_pct_target.toFixed(0)+'%)',n(pp.naat)+' ('+pp.naat_pct_target.toFixed(0)+'%)',chgN(pd.naat,pp.naat,true),pd.naat_pct_target.toFixed(0)+'%','n/a',{t:'n/a',c:MUT}),
 row('X-ray offered (% of enrolled)',pct(pd.xray_offered,pd.enrolled),pct(pp.xray_offered,pp.enrolled),chgPP(pd.xray_offered_pct,pp.xray_offered_pct,true),'–','n/a',{t:'n/a',c:MUT}),
 row('Notified – diagnosing base (n)',n(nd.notified),n(np.notified),chgN(nd.notified,np.notified,true),nd.pct_target.toFixed(0)+'%',n(nl.notified),chgN(nd.notified,nl.notified,true)),
 row('Notified – enrolment base (n)',n(N.enrol.last.District.notified),n(N.enrol.prev.District.notified),chgN(N.enrol.last.District.notified,N.enrol.prev.District.notified,true),N.enrol.last.District.pct_target.toFixed(0)+'%',n(N.enrol.lastyear.District.notified),chgN(N.enrol.last.District.notified,N.enrol.lastyear.District.notified,true)),
 row('Private notified (n, diag. base)',n(nd.private),n(np.private),chgN(nd.private,np.private,true),'–',n(nl.private),chgN(nd.private,nl.private,true)),
 row('Treatment started >7 days (% of started)',tl.late_pct_of_started.toFixed(1)+'% ('+tl.late_gt7+'/'+tl.started+')',tp.late_pct_of_started.toFixed(1)+'%',chgPP(tl.late_pct_of_started,tp.late_pct_of_started,false),'–',tly.late_pct_of_started.toFixed(1)+'%',chgPP(tl.late_pct_of_started,tly.late_pct_of_started,false)),
 row('Not initiated (n)',n(tl.not_initiated),n(tp.not_initiated),chgN(tl.not_initiated,tp.not_initiated,false),'–',n(tly.not_initiated)+'*',{t:'see notes',c:MUT}),
 row('Outcome success – '+lab(O.cohort_months.last)+' cohort',ol.success_pct.toFixed(1)+'% ('+ol.success+'/'+ol.denominator+')',op.success_pct.toFixed(1)+'% ('+lab(O.cohort_months.prev)+')',chgPP(ol.success_pct,op.success_pct,true),'–','n/a',{t:'n/a',c:MUT}),
 row('Died – '+lab(O.cohort_months.last)+' cohort',ol.died_pct.toFixed(1)+'% ('+ol.died+'/'+ol.denominator+')',op.died_pct.toFixed(1)+'%',chgPP(ol.died_pct,op.died_pct,false),'–','n/a',{t:'n/a',c:MUT}),
 row('PTFU done (done/due)',ptL[3]+'% ('+n(ptL[2])+'/'+n(ptL[1])+')',ptP[3]+'% ('+n(ptP[2])+'/'+n(ptP[1])+')',chgPP(ptL[3],ptP[3],true),'–','n/a',{t:'n/a',c:MUT}),
 row('DTC done (done/due)',pct(dL[0],dL[1],0)+' ('+n(dL[0])+'/'+n(dL[1])+')',pct(dP[0],dP[1],0)+' ('+n(dP[0])+'/'+n(dP[1])+')',chgPP(100*dL[0]/dL[1],100*dP[0]/dP[1],true),'–','n/a',{t:'n/a',c:MUT})];
 tbl(s,['Indicator (district)',L,PV,'Change','Target %',LY,'Change vs '+LY],rows,{colW:[2.9,1.4,1.3,1.0,.8,.9,.9],hh:.3,rh:.275,fs:8.5,y:.88});}

// 3 PRESUMPTIVE
{const mk=marks(T.map(t=>P.last[t].tested_pct_target),T.map(t=>P.last[t].enrolled),true);
 const s=base('1 · Case-finding','Presumptive: enrolled and tested vs monthly target','Tested reached '+pd.tested_pct_target.toFixed(0)+'% of target only because enrolment jumped '+Math.round(100*(pd.enrolled-pp.enrolled)/pp.enrolled)+'% to '+n(pd.enrolled)+'; the share of enrolled who were tested stayed at '+pd.tested_pct_enrolled.toFixed(0)+'% ('+pp.tested_pct_enrolled.toFixed(0)+'% in '+PV+'). '+LY+': n/a – the register holds no 2025 enrolments.');
 const rows=[...T,'District'].map(t=>{const a=P.last[t],b=P.prev[t];return [t==='District'?'District':mk(t),n(Math.round(a.target)),cell(n(a.tested),a.tested_pct_target.toFixed(0)+'% of target'),cell(n(b.tested),b.tested_pct_target.toFixed(0)+'% of target'),chgPP(a.tested_pct_target,b.tested_pct_target,true),cell(n(a.enrolled),'enrolled'),cell(n(b.enrolled),'enrolled'),pct(a.tested,a.enrolled,0)+' / '+pct(b.tested,b.enrolled,0)]});
 tbl(s,['Taluka','Monthly target','Tested '+L+' (n, %)','Tested '+PV+' (n, %)','Change in % target','Enrolled '+L,'Enrolled '+PV,'Tested ÷ enrolled ('+L.slice(0,3)+' / '+PV.slice(0,3)+')'],rows,{colW:[1.35,.8,1.25,1.25,1.1,1.0,1.0,1.45],totalRow:true,rh:.38,fs:9});
 s.addText(leg,{x:.4,y:4.5,w:9.2,h:.22,fontSize:8,color:MUT,margin:0,isTextBox:true});}

// 4 NAAT / MICRO
{const mk=marks(T.map(t=>P.last[t].naat_pct_target),T.map(t=>P.last[t].enrolled),true);
 const s=base('1 · Case-finding','NAAT vs microscopy','NAAT was '+pd.naat_pct_target.toFixed(0)+'% of target ('+n(pd.naat)+' vs '+n(pp.naat)+' in '+PV+') while microscopy rose from '+n(pp.micro)+' to '+n(pd.micro)+'. NAAT share of tests fell from '+naiP.toFixed(0)+'% to '+nai.toFixed(0)+'% – a >50% swing: verify cartridge stock/machine uptime and result-date entry before drawing conclusions.');
 const rows=[...T,'District'].map(t=>{const a=P.last[t],b=P.prev[t];return [t==='District'?'District':mk(t),cell(n(a.naat),a.naat_pct_target.toFixed(0)+'% of target'),cell(n(b.naat),b.naat_pct_target.toFixed(0)+'% of target'),chgPP(a.naat_pct_target,b.naat_pct_target,true),n(a.micro),n(b.micro),pct(a.naat,a.tested,0)+' / '+pct(b.naat,b.tested,0)]});
 tbl(s,['Taluka','NAAT '+L+' (n, %)','NAAT '+PV+' (n, %)','Change in % target','Microscopy '+L,'Microscopy '+PV,'NAAT share of tests ('+L.slice(0,3)+' / '+PV.slice(0,3)+')'],rows,{colW:[1.4,1.4,1.4,1.2,1.1,1.1,1.6],totalRow:true,rh:.38,fs:9.5});
 s.addText(leg+' '+LY+': n/a.',{x:.4,y:4.5,w:9.2,h:.22,fontSize:8,color:MUT,margin:0,isTextBox:true});}

// 5 XRAY
{const mk=marks(T.map(t=>P.last[t].xray_offered_pct),T.map(t=>P.last[t].enrolled),true);
 const s=base('1 · Case-finding','X-ray offered and abnormal-X-ray follow-through','X-ray offered fell to '+pd.xray_offered_pct.toFixed(1)+'% ('+pp.xray_offered_pct.toFixed(1)+'% in '+PV+') as enrolment surged. Abnormal X-ray follow-through is small numbers ('+pd.abnormal_tested+' of '+pd.abnormal_xray+' tested) – the issue is that X-ray results are rarely entered. Half of those tested are recorded “Asymptomatic”.');
 const rows=[...T,'District'].map(t=>{const a=P.last[t],b=P.prev[t];return [t==='District'?'District':mk(t),cell(n(a.xray_offered)+'/'+n(a.enrolled),pct(a.xray_offered,a.enrolled)),cell(n(b.xray_offered)+'/'+n(b.enrolled),pct(b.xray_offered,b.enrolled)),chgPP(a.xray_offered_pct,b.xray_offered_pct,true),a.abnormal_xray?a.abnormal_tested+'/'+a.abnormal_xray:'0/0',cell(n(a.asym_tested),pct(a.asym_tested,a.tested,0)+' of tested'),cell(n(b.asym_tested),pct(b.asym_tested,b.tested,0)+' of tested')]});
 tbl(s,['Taluka','Offered '+L+' (n/N, %)','Offered '+PV+' (n/N, %)','Change','Abnormal X-ray tested '+L,'Asymptomatic-labelled but tested '+L,'Same, '+PV],rows,{colW:[1.4,1.4,1.4,1.1,1.2,1.4,1.3],totalRow:true,rh:.38,fs:9.5});
 s.addText(leg+' Lower “asymptomatic-labelled” is better. '+LY+': n/a.',{x:.4,y:4.5,w:9.2,h:.22,fontSize:8,color:MUT,margin:0,isTextBox:true});}

// 6 NOTIFICATION
{const mk=marks(T.map(t=>N.diag.last[t].pct_target),T.map(t=>N.diag.last[t].monthly_target),true);
 const s=base('2 · Notification','Notification vs monthly target – diagnosing and enrolment base','Notified '+nd.notified+' ('+nd.pct_target.toFixed(0)+'% of the monthly target of '+Math.round(nd.monthly_target)+'): '+Math.abs(Math.round(100*(nd.notified-np.notified)/np.notified))+'% below '+PV+' and '+Math.abs(Math.round(100*(nd.notified-nl.notified)/nl.notified))+'% below '+LY+'. Lunawada fell to '+N.diag.last.Lunawada.notified+' (from '+N.diag.prev.Lunawada.notified+' and '+N.diag.lastyear.Lunawada.notified+'); Santrampur matches last year.');
 const rows=[...T,'District'].map(t=>{const a=N.diag.last[t],b=N.diag.prev[t],c=N.diag.lastyear[t],e=N.enrol.last[t];return [t==='District'?'District':mk(t),a.monthly_target.toFixed(0),cell(a.notified,a.pct_target.toFixed(0)+'%'),cell(b.notified,b.pct_target.toFixed(0)+'%'),cell(c.notified,c.pct_target.toFixed(0)+'%'),chgN(a.notified,b.notified,true),chgN(a.notified,c.notified,true),cell(e.notified,e.pct_target.toFixed(0)+'%')]});
 tbl(s,['Taluka','Monthly target','Diag. base '+L,'Diag. base '+PV,'Diag. base '+LY,'vs '+PV,'vs '+LY,'Enrol. base '+L],rows,{colW:[1.35,.8,1.2,1.2,1.2,1.1,1.1,1.25],totalRow:true,rh:.38,fs:9.5});
 s.addText(leg+' Cells: n, then % of monthly target.',{x:.4,y:4.5,w:9.2,h:.22,fontSize:8,color:MUT,margin:0,isTextBox:true});}

// 7 PUBLIC / PRIVATE
{const TG=R.ntgt_pubpriv;const pubT=Object.values(TG).reduce((a,b)=>a+b[0],0),prvT=Object.values(TG).reduce((a,b)=>a+b[1],0);
 const mk=marks(T.map(t=>N.diag.last[t].pct_target),T.map(t=>N.diag.last[t].monthly_target),true);
 const s=base('2 · Notification','Public vs private notification (diagnosing base)','Public notified '+nd.public+' ('+Math.round(100*nd.public/pubT)+'% of its monthly target of '+Math.round(pubT)+') vs '+np.public+' in '+PV+'; private '+nd.private+' ('+Math.round(100*nd.private/prvT)+'% of '+Math.round(prvT)+') vs '+np.private+' and '+nl.private+' in '+LY+'. Private remains less than half of last year.');
 const rows=[...T,'District'].map(t=>{const a=N.diag.last[t],b=N.diag.prev[t],c=N.diag.lastyear[t];const pt=t==='District'?pubT:TG[t][0],vt=t==='District'?prvT:TG[t][1];
  return [t==='District'?'District':mk(t),cell(a.public,pt?Math.round(100*a.public/pt)+'% of '+Math.round(pt):'–'),b.public,c.public,chgN(a.public,c.public,true),cell(a.private,vt?Math.round(100*a.private/vt)+'% of '+Math.round(vt):'no target'),b.private,c.private,chgN(a.private,c.private,true)]});
 tbl(s,['Taluka','Public '+L,'Public '+PV,'Public '+LY,'vs '+LY,'Private '+L,'Private '+PV,'Private '+LY,'vs '+LY],rows,{colW:[1.35,1.3,.8,.8,.95,1.3,.8,.8,1.1],totalRow:true,rh:.38,fs:9.5});
 s.addText(leg+' Targets: TU-wise public/private monthly targets.',{x:.4,y:4.5,w:9.2,h:.22,fontSize:8,color:MUT,margin:0,isTextBox:true});}

// 8 TREATMENT START
{const mk=marks(T.map(t=>TS.last[t].late_pct_of_started),T.map(t=>TS.last[t].started),false);
 const s=base('3 · Treatment initiation','Treatment start within 7 days and not initiated','Late starts improved to '+tl.late_pct_of_started.toFixed(1)+'% ('+tp.late_pct_of_started.toFixed(1)+'% in '+PV+', '+tly.late_pct_of_started.toFixed(1)+'% in '+LY+'), but '+tl.not_initiated+' of '+tl.diagnosed+' diagnosed patients have no start date – Lunawada '+TS.last.Lunawada.not_initiated+' and Santrampur '+TS.last.Santrampur.not_initiated+'. Trace them now.');
 const f=(a)=>cell(a.late_gt7+'/'+a.started,a.late_pct_of_started.toFixed(1)+'%');
 const rows=[...T,'District'].map(t=>{const a=TS.last[t],b=TS.prev[t],c=TS.lastyear[t];return [t==='District'?'District':mk(t),f(a),f(b),f(c),chgPP(a.late_pct_of_started,b.late_pct_of_started,false),cell(a.not_initiated,'of '+a.diagnosed),cell(b.not_initiated,'of '+b.diagnosed),cell(c.not_initiated+'*','of '+c.diagnosed)]});
 tbl(s,['Taluka (current TBU)','Started >7 d '+L+' (n/N, %)','Started >7 d '+PV,'Started >7 d '+LY,'Change vs '+PV,'Not initiated '+L,'Not initiated '+PV,'Not initiated '+LY],rows,{colW:[1.4,1.35,1.2,1.2,1.1,1.0,1.0,.95],totalRow:true,rh:.38,fs:9});
 s.addText(leg+' Lower is better. *Last-year not-initiated = 0 should be read cautiously (see notes).',{x:.4,y:4.5,w:9.2,h:.22,fontSize:8,color:MUT,margin:0,isTextBox:true});}

// 9 OUTCOME
{const mk=marks(T.map(t=>O.last[t].success_pct),T.map(t=>O.last[t].denominator),true);const cl=lab(O.cohort_months.last),cp=lab(O.cohort_months.prev);
 const s=base('4 · Treatment outcome','Outcome of the matured cohort – patients notified in '+cl,'This is the '+cl+' notification cohort (9-month rule), not '+L+' activity. Success '+ol.success_pct.toFixed(1)+'% vs '+op.success_pct.toFixed(1)+'% for '+cp+'; deaths improved to '+ol.died_pct.toFixed(1)+'% from '+op.died_pct.toFixed(1)+'%. Virpur ('+O.last.Virpur.died+' of '+O.last.Virpur.denominator+') and Santrampur ('+O.last.Santrampur.died+' of '+O.last.Santrampur.denominator+') deaths are small numbers; lowest success: '+(()=>{const w=T.map(t=>[t,O.last[t].success_pct,O.last[t].denominator]).filter(x=>x[2]>=10).sort((a,b)=>a[1]-b[1])[0];return w[0]+' '+w[1].toFixed(0)+'%'})()+'.');
 const rows=[...T,'District'].map(t=>{const a=O.last[t],b=O.prev[t];return [t==='District'?'District':mk(t),a.denominator,cell(a.success+'/'+a.denominator,a.success_pct.toFixed(1)+'%'),cell(b.success+'/'+b.denominator,b.success_pct.toFixed(1)+'%'),chgPP(a.success_pct,b.success_pct,true),cell(a.died+'/'+a.denominator,a.died_pct.toFixed(1)+'%'),cell(b.died+'/'+b.denominator,b.died_pct.toFixed(1)+'%'),a.pending+' ('+a.pending_pct.toFixed(1)+'%)']});
 tbl(s,['Taluka','Notified '+cl,'Success '+cl+' (n/N, %)','Success '+cp+' (n/N, %)','Change','Died '+cl+' (n/N, %)','Died '+cp+' (n/N, %)','Pending '+cl],rows,{colW:[1.3,.8,1.35,1.35,.95,1.3,1.3,.85],totalRow:true,rh:.38,fs:9});
 s.addText(leg+' DS-TB 2HRZE/4HRE pulmonary. '+lab(O.cohort_months.lastyear)+' cohort (same month last year): n/a – not in the register.',{x:.4,y:4.5,w:9.2,h:.22,fontSize:8,color:MUT,margin:0,isTextBox:true});}

// 10 PTFU + DTC
{const sd=t=>R.dtc2.data.Sep[t].reduce((a,b)=>[a[0]+b[0],a[1]+b[1]],[0,0]),ad=t=>R.dtc2.data.Aug[t].reduce((a,b)=>[a[0]+b[0],a[1]+b[1]],[0,0]);
 const mk=marks(T.map(t=>R.ptfu[t][3][3]),T.map(t=>R.ptfu[t][3][1]),true);
 const s=base('5 · Follow-up','Post-treatment follow-up (PTFU) and differentiated care (DTC)','Both slipped: PTFU '+ptL[3]+'% vs '+ptP[3]+'% and DTC '+Math.round(100*dL[0]/dL[1])+'% vs '+Math.round(100*dP[0]/dP[1])+'%. Lunawada is lowest on both (PTFU '+R.ptfu.Lunawada[3][3]+'%, DTC '+Math.round(100*sd('Lunawada')[0]/sd('Lunawada')[1])+'%); Khanpur leads. Sep cohorts are still partly inside their windows.');
 const pc2=a=>cell(n(a[2])+'/'+n(a[1]),a[3]+'%');
 const rows=[...T,'District'].map(t=>{const a=R.ptfu[t][3],b=R.ptfu[t][2],x=sd(t),y=ad(t);return [t==='District'?'District':mk(t),pc2(a),pc2(b),chgPP(a[3],b[3],true),cell(n(x[0])+'/'+n(x[1]),Math.round(100*x[0]/x[1])+'%'),cell(n(y[0])+'/'+n(y[1]),Math.round(100*y[0]/y[1])+'%'),chgPP(100*x[0]/x[1],100*y[0]/y[1],true)]});
 tbl(s,['Taluka','PTFU '+L+' (done/due, %)','PTFU '+PV,'Change','DTC '+L+' (done/due, %)','DTC '+PV,'Change'],rows,{colW:[1.4,1.45,1.35,1.0,1.5,1.35,1.15],totalRow:true,rh:.38,fs:9.5});
 s.addText(leg+' '+LY+': n/a – no 2025 follow-up data supplied.',{x:.4,y:4.5,w:9.2,h:.22,fontSize:8,color:MUT,margin:0,isTextBox:true});}

// 11 IP / CP / TPT
{const U={Balasinor:'Balasinor',Khanpur:'Khanpur',Lunawada:'Lunawada',Santrampur:'Santrampur',Virpur:'Virpur',District:'District'};
 const ip=R.ip2,cp=R.cp2;const ii=ip.months.indexOf('2026-07'),ij=ip.months.indexOf('2026-06'),ci=cp.months.indexOf('2026-03'),cj=cp.months.indexOf('2026-02');
 const fc=v=>v&&v[0]?cell(v[1]+'/'+v[0],Math.round(100*v[1]/v[0])+'%'):'–';
 const tm=(t,i)=>R.tpt.tal_month[t]?R.tpt.tal_month[t][i]:R.tpt.monthly[i];
 const tg=v=>v&&v[2]?cell(v[3]+'/'+v[2],Math.round(100*v[3]/v[2])+'%'):'–';
 const s=base('5 · Follow-up','IP, CP and TPT – latest fully due cohorts','IP: '+lab('2026-07')+' cohort '+Math.round(100*ip.units.District[ii][1]/ip.units.District[ii][0])+'% vs '+Math.round(100*ip.units.District[ij][1]/ip.units.District[ij][0])+'% for '+lab('2026-06')+'. CP: '+lab('2026-03')+' cohort '+Math.round(100*cp.units.District[ci][1]/cp.units.District[ci][0])+'% vs '+Math.round(100*cp.units.District[cj][1]/cp.units.District[cj][0])+'%. TPT (Aug, latest month with most contacts traced): '+Math.round(100*tm('District',7)[3]/tm('District',7)[2])+'% of contacts given TPT. Cohort-based, so shown for the latest due cohorts, not '+L+'.');
 const rows=[...T,'District'].map(t=>{const u=t==='Kadana'?null:U[t];const a=t==='District'?R.tpt.monthly[7]:R.tpt.tal_month[t][7],b=t==='District'?R.tpt.monthly[6]:R.tpt.tal_month[t][6];
  return [t,u?fc(ip.units[u][ii]):'in Lunawada',u?fc(ip.units[u][ij]):'in Lunawada',u?fc(cp.units[u][ci]):'in Lunawada',u?fc(cp.units[u][cj]):'in Lunawada',tg(a),tg(b)]});
 tbl(s,['Taluka','IP – '+lab('2026-07')+' cohort (done/due)','IP – '+lab('2026-06')+' cohort','CP – '+lab('2026-03')+' cohort','CP – '+lab('2026-02')+' cohort','TPT given Aug (given/contacts)','TPT given Jul'],rows,{colW:[1.3,1.4,1.3,1.4,1.3,1.5,1.0],totalRow:true,rh:.38,fs:9.5});
 s.addText('Cell: done/due (or given/contacts), then %. Kadana (non-functional TBU) is reported under Lunawada for IP/CP. '+LY+': n/a. '+L+' TPT is partial (contacts still being traced).',{x:.4,y:4.5,w:9.2,h:.22,fontSize:8,color:MUT,margin:0,isTextBox:true});}

// 12 IMPROVED / SLIPPED / ACTIONS
{const s=base('Summary','What improved, what slipped, and 5 actions for next month',null);
 const col=(x,w,h,c,items)=>{s.addShape(p.ShapeType.roundRect,{x,y:.92,w,h:3.9,fill:{color:'FFFFFF'},rectRadius:.08,line:{color:c,width:1.5}});
  s.addText(h,{x:x+.12,y:.98,w:w-.2,h:.3,fontFace:'Cambria',fontSize:13,bold:true,color:c,margin:0,isTextBox:true});
  s.addText(items.map((t,i)=>({text:t,options:{bullet:true,breakLine:i<items.length-1}})),{x:x+.12,y:1.34,w:w-.24,h:3.4,fontSize:9.5,color:INK,paraSpaceAfter:5,margin:0,isTextBox:true,valign:'top'});};
 col(.4,2.9,'Improved',GOOD,['Treatment started >7 days: '+tl.late_pct_of_started.toFixed(1)+'% (was '+tp.late_pct_of_started.toFixed(1)+'%; '+tly.late_pct_of_started.toFixed(1)+'% in '+LY+')','Deaths in the '+lab(O.cohort_months.last)+' cohort '+ol.died_pct.toFixed(1)+'% (was '+op.died_pct.toFixed(1)+'%)','Presumptive enrolment +'+Math.round(100*(pd.enrolled-pp.enrolled)/pp.enrolled)+'% to '+n(pd.enrolled),'Santrampur notified '+N.diag.last.Santrampur.notified+' (+'+Math.round(100*(N.diag.last.Santrampur.notified-N.diag.prev.Santrampur.notified)/N.diag.prev.Santrampur.notified)+'% vs '+PV+(N.diag.last.Santrampur.notified===N.diag.lastyear.Santrampur.notified?', same as '+LY:'; '+LY+': '+N.diag.lastyear.Santrampur.notified)+')','Private notified '+nd.private+' vs '+np.private+' in '+PV]);
 col(3.45,2.9,'Slipped',BAD,['NAAT '+n(pd.naat)+' vs '+n(pp.naat)+' (−'+Math.round(100*(pp.naat-pd.naat)/pp.naat)+'%); microscopy '+n(pd.micro)+' vs '+n(pp.micro),'X-ray offered '+pd.xray_offered_pct.toFixed(1)+'% (was '+pp.xray_offered_pct.toFixed(1)+'%)','Notified '+nd.notified+' vs '+np.notified+' ('+PV+') and '+nl.notified+' ('+LY+')','Lunawada notified '+N.diag.last.Lunawada.notified+' (from '+N.diag.prev.Lunawada.notified+'); Kadana '+N.diag.last.Kadana.notified+' (from '+N.diag.prev.Kadana.notified+')','PTFU '+ptL[3]+'% (was '+ptP[3]+'%); DTC '+Math.round(100*dL[0]/dL[1])+'% (was '+Math.round(100*dP[0]/dP[1])+'%)','Not initiated '+tl.not_initiated+' (was '+tp.not_initiated+')']);
 col(6.5,3.1,'5 actions for next month',NAVY,['Audit NAAT cartridge stock and machine uptime this week; restore NAAT-first testing','Test every enrolled presumptive: raise tested ÷ enrolled from '+pd.tested_pct_enrolled.toFixed(0)+'%; enter every X-ray result','Lunawada and Kadana notification drive; private-sector (PPSA) engagement to recover to '+LY+' levels','Trace the '+tl.not_initiated+' patients with no start date (Lunawada '+TS.last.Lunawada.not_initiated+', Santrampur '+TS.last.Santrampur.not_initiated+')','Clear PTFU/DTC due-lists in Lunawada before month-end']);}

// 13 NOTES
{const s=base('Notes','Data notes, assumptions and n/a reasons',null);
 s.addText([{text:'Method',options:{bold:true,color:NAVY,breakLine:true}},
 {text:'Presumptive “tested” is counted by the month of ENROLMENT (monthly-review rule): '+L+' = '+n(pd.tested)+'. The yearly district deck counts by the month of test date (Sep = '+n(R.pres['Total (District)'].a[8])+'), so the two differ by design.',options:{bullet:true,breakLine:true}},
 {text:'Notification registers (diagnosing and enrolment base) were regenerated from NOTIFICATION.xlsx extracted '+AS+'; the older Notification_Register files in the uploads (2-Oct) were not used.',options:{bullet:true,breakLine:true}},
 {text:'Targets: presumptive/NAAT monthly CBNAAT targets; notification annual ÷ 12; public/private from the TU-wise 2026 targets. No RAG thresholds were invented – arrows show direction only.',options:{bullet:true,breakLine:true}},
 {text:'Outcome uses the 9-month rule: '+L+' review shows the '+lab(O.cohort_months.last)+' notification cohort vs '+lab(O.cohort_months.prev)+'.',options:{bullet:true,breakLine:true}},
 {text:'n/a and cautions',options:{bold:true,color:NAVY,breakLine:true}},
 {text:LY+' is n/a for presumptive/NAAT/X-ray (register holds no 2025 enrolments), outcome (Dec-24 cohort not in the register) and all follow-up indicators (no 2025 data supplied).',options:{bullet:true,breakLine:true}},
 {text:'*Not initiated for '+LY+' shows 0 in the current register; older patients are mostly closed out, so do not read this as a true zero.',options:{bullet:true,breakLine:true}},
 {text:'NAAT/microscopy swing (NAAT '+Math.round(100*(pd.naat-pp.naat)/pp.naat)+'%, microscopy '+Math.round(pd.micro/pp.micro)+'×) is >50% – verify against lab registers. Abnormal X-ray follow-through ('+pd.abnormal_tested+'/'+pd.abnormal_xray+') is small numbers.',options:{bullet:true,breakLine:true}},
 {text:'IP/CP/TPT are cohort-based (diagnosis month), so latest fully due cohorts are shown. PTFU uses the cumulative file as uploaded (check its extract date). DTC rule: due = diagnosis in the 28-day window for that visit; done = matching assessment recorded.',options:{bullet:true}}],{x:.5,y:.95,w:9,h:4.2,fontSize:10.5,color:INK,paraSpaceAfter:4,margin:0,isTextBox:true,valign:'top'});}

p.writeFile({fileName:OUTF}).then(()=>console.log('written',sl));

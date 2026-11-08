/* СППР «Солнечная Энергетика РФ» — вся логика v6.1 */
(function(){
"use strict";

/* ---------- ХЕЛПЕРЫ ---------- */
function $(id){ return document.getElementById(id); }
function qsa(s){ return Array.prototype.slice.call(document.querySelectorAll(s)); }
function on(el,ev,fn){ if(el) el.addEventListener(ev,fn); }
function plotlyReady(){ return typeof window.Plotly !== "undefined"; }
function xlsxReady(){ return typeof window.XLSX !== "undefined"; }
function docxReady(){ return typeof window.docx !== "undefined"; }
function jszipReady(){ return typeof window.JSZip !== "undefined"; }

var REGIONS = [
  "Астраханская обл.","Волгоградская обл.","Оренбургская обл.","Самарская обл.","Саратовская обл.",
  "Республика Башкортостан","Республика Алтай","Республика Крым","Ставропольский край",
  "Краснодарский край","Ростовская обл.","Белгородская обл.","Республика Дагестан",
  "Чеченская Республика","Республика Бурятия","Забайкальский край","Приморский край"
];
var STATUSES = ["факт","оценка","прогноз"];
var PALETTE = ["#f59e0b","#1e40af","#0891b2","#7c3aed","#dc2626","#059669","#a16207","#0f766e","#9333ea","#b91c1c","#0891b2","#7c3aed","#d97706","#4f46e5","#10b981","#f472b6","#8b5cf6"];
var CONSENT_KEY = "solar_rf_consent_v6";

var state = {
  data: [], selected: {}, theme: "light", expMode: "short",
  optUnit: "mln", optPrec: 2, optSep: "ru",
  consent: null, currentPage: "home",
  econ: {}
};

/* дефолтные параметры экономики — используются и при сбросе */
var ECON_DEFAULTS = {
  region: "",
  power: 10,            // МВт
  yieldPerKW: 1400,     // кВт·ч / кВт·год
  capex: 90,            // млн ₽
  opexPct: 1.5,         // % CAPEX/год
  tariff: 4.5,          // ₽/кВт·ч
  degrType: "linear",   // linear | exp
  degr: 0.5,            // %/год
  discount: 12,         // %
  life: 25,             // лет
  tariffGrowth: 5,      // %/год
  opexGrowth: 4,        // %/год
  tax: 20,              // % (налог на прибыль РФ)
  subsidy: 0,           // млн ₽/год (ДПМ ВИЭ, зелёные сертификаты)
  waccMode: false,
  debtPct: 70, debtRate: 10, equityRate: 18 // для WACC
};

var LOG = { info: [], err: [], warn: [] };
function tlog(level,msg){
  var e = { t:new Date(), level:level, msg:msg };
  if(level==="ERR") LOG.err.push(e); else if(level==="WARN") LOG.warn.push(e);
  LOG.info.push(e); renderTerm();
}
function tErr(m){ tlog("ERR",m); console.error(m); }
function tWarn(m){ tlog("WARN",m); console.warn(m); }
function tInfo(m){ tlog("INFO",m); }

function escapeHtml(s){
  return String(s==null?"":s).replace(/[&<>"']/g,function(c){
    return {"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c];
  });
}
function toast(text,kind){
  var t = document.createElement("div");
  t.className = "toast "+(kind||"ok"); t.textContent = text;
  $("toasts").appendChild(t);
  setTimeout(function(){t.classList.add("show");},10);
  setTimeout(function(){t.classList.remove("show");setTimeout(function(){t.remove();},300);},3500);
}
function showMsg(id,text,kind){ var el=$(id); if(!el) return; el.className="msg show "+(kind||"ok"); el.textContent=text; }
function fmtDate(d){
  if(!d) return "—";
  var dt = new Date(d); if(isNaN(dt.getTime())) return d;
  return String(dt.getDate()).padStart(2,"0")+"."+String(dt.getMonth()+1).padStart(2,"0")+"."+dt.getFullYear();
}
function fmtNum(x,d){
  if(x===null||x===undefined||isNaN(x)||!isFinite(x)) return "—";
  d = (typeof d==="number")?d:state.optPrec;
  return Number(x).toLocaleString(state.optSep==="en"?"en-US":"ru-RU",{minimumFractionDigits:d,maximumFractionDigits:d});
}
function fmtPct(x,d){ if(x==null||!isFinite(x)) return "—"; d = (d==null?2:d); return fmtNum(x,d)+"%"; }
function fmtMoney(x,d){ if(x==null||!isFinite(x)) return "—"; return fmtNum(x, d==null?2:d)+" млн ₽"; }
function unitLabel(){ return state.optUnit==="bln"?"млрд кВт·ч":state.optUnit==="twh"?"ТВт·ч":"млн кВт·ч"; }
function convertUnits(v){
  if(v==null) return v;
  if(state.optUnit==="bln") return v/1000;
  if(state.optUnit==="twh") return v/1e6;
  return v;
}

/* ---------- МАТЕМАТИКА ---------- */
function mean(a){ if(!a.length) return null; var s=0; for(var i=0;i<a.length;i++) s+=a[i]; return s/a.length; }
function median(a){
  if(!a.length) return null;
  var s = a.slice().sort(function(x,y){return x-y;});
  var m = Math.floor(s.length/2);
  return s.length%2 ? s[m] : (s[m-1]+s[m])/2;
}
function std(a){
  if(a.length<2) return null;
  var m=mean(a), s=0;
  for(var i=0;i<a.length;i++) s+=(a[i]-m)*(a[i]-m);
  return Math.sqrt(s/(a.length-1));
}
function sum(a){ var s=0; for(var i=0;i<a.length;i++) s+=a[i]; return s; }
function minOf(a){ var m=null; for(var i=0;i<a.length;i++) if(m===null||a[i]<m) m=a[i]; return m; }
function maxOf(a){ var m=null; for(var i=0;i<a.length;i++) if(m===null||a[i]>m) m=a[i]; return m; }
function linreg(y){
  var n=y.length; if(n<2) return null;
  var tm=(n-1)/2, ym=mean(y), num=0, den=0;
  for(var i=0;i<n;i++){ num+=(i-tm)*(y[i]-ym); den+=(i-tm)*(i-tm); }
  var b = den===0?0:num/den, a = ym-b*tm;
  var ssTot=0, ssRes=0;
  for(var j=0;j<n;j++){ ssTot+=(y[j]-ym)*(y[j]-ym); var pr=a+b*j; ssRes+=(y[j]-pr)*(y[j]-pr); }
  return {a:a,b:b,r2: ssTot===0?1:1-ssRes/ssTot};
}
function polyreg(y,degree){
  var n=y.length; if(n<degree+1) return null;
  var X=[],Y=[];
  for(var i=0;i<n;i++){ var r=[]; for(var d=0;d<=degree;d++) r.push(Math.pow(i,d)); X.push(r); Y.push(y[i]); }
  var dim=degree+1, A=[],B=[];
  for(var rr=0;rr<dim;rr++){
    A.push([]);
    for(var cc=0;cc<dim;cc++){ var s=0; for(var k=0;k<n;k++) s+=X[k][rr]*X[k][cc]; A[rr].push(s); }
    var s2=0; for(var k2=0;k2<n;k2++) s2+=X[k2][rr]*Y[k2]; B.push(s2);
  }
  var M = A.map(function(row,i){ return row.concat([B[i]]); });
  for(var i2=0;i2<dim;i2++){
    var piv=i2;
    for(var r2=i2+1;r2<dim;r2++) if(Math.abs(M[r2][i2])>Math.abs(M[piv][i2])) piv=r2;
    if(Math.abs(M[piv][i2])<1e-12) return null;
    var tmp=M[i2]; M[i2]=M[piv]; M[piv]=tmp;
    for(var r3=i2+1;r3<dim;r3++){ var f=M[r3][i2]/M[i2][i2]; for(var c3=i2;c3<=dim;c3++) M[r3][c3]-=f*M[i2][c3]; }
  }
  var coef=new Array(dim).fill(0);
  for(var i3=dim-1;i3>=0;i3--){ var s3=M[i3][dim]; for(var j3=i3+1;j3<dim;j3++) s3-=M[i3][j3]*coef[j3]; coef[i3]=s3/M[i3][i3]; }
  var ym=mean(y), ssTot=0, ssRes=0;
  for(var i4=0;i4<n;i4++){
    ssTot+=(y[i4]-ym)*(y[i4]-ym);
    var pr=0; for(var d2=0;d2<=degree;d2++) pr+=coef[d2]*Math.pow(i4,d2);
    ssRes+=(y[i4]-pr)*(y[i4]-pr);
  }
  return {coef:coef, r2: ssTot===0?1:1-ssRes/ssTot,
    predict:function(t){ var v=0; for(var d3=0;d3<=degree;d3++) v+=coef[d3]*Math.pow(t,d3); return v; }};
}
function expreg(y){
  if(y.length<2) return null;
  var ys=y.map(function(v){ return v>0?Math.log(v):null; });
  if(ys.some(function(v){return v===null;})) return null;
  var lr=linreg(ys); if(!lr) return null;
  var ym=mean(y), ssTot=0, ssRes=0;
  for(var i=0;i<y.length;i++){ ssTot+=(y[i]-ym)*(y[i]-ym); var pr=Math.exp(lr.a+lr.b*i); ssRes+=(y[i]-pr)*(y[i]-pr); }
  return {a:Math.exp(lr.a), b:lr.b, r2: ssTot===0?1:1-ssRes/ssTot,
    predict:function(t){ return Math.exp(lr.a+lr.b*t); }};
}
function mape(y,fit){
  var n=Math.min(y.length,fit.length), s=0, c=0;
  for(var i=0;i<n;i++) if(y[i]!==0){ s+=Math.abs((y[i]-fit[i])/y[i]); c++; }
  return c?s/c*100:null;
}
function cagrSeries(years,vals){
  if(vals.length<2) return null;
  var i0=0,i1=0;
  for(var i=1;i<years.length;i++){ if(years[i]<years[i0]) i0=i; if(years[i]>years[i1]) i1=i; }
  var n=years[i1]-years[i0], v0=vals[i0], v1=vals[i1];
  if(n<=0||v0<=0||v1<=0) return null;
  return (Math.pow(v1/v0,1/n)-1)*100;
}
function pearson(x,y){
  var n=x.length; if(n<2||y.length!==n) return null;
  var mx=mean(x), my=mean(y), num=0, dx=0, dy=0;
  for(var i=0;i<n;i++){ num+=(x[i]-mx)*(y[i]-my); dx+=(x[i]-mx)*(x[i]-mx); dy+=(y[i]-my)*(y[i]-my); }
  if(dx===0||dy===0) return null;
  return num/Math.sqrt(dx*dy);
}
function movingAvg(v,w){
  w=w||3; var o=[];
  for(var i=0;i<v.length;i++){ var from=Math.max(0,i-w+1); var sl=v.slice(from,i+1); o.push(sum(sl)/sl.length); }
  return o;
}
function iqrMethod(v){
  if(v.length<4) return {lo:null,hi:null};
  var s=v.slice().sort(function(a,b){return a-b;});
  function q(p){ var idx=(s.length-1)*p, lo=Math.floor(idx), hi=Math.ceil(idx); if(lo===hi) return s[lo]; return s[lo]+(s[hi]-s[lo])*(idx-lo); }
  var q1=q(0.25), q3=q(0.75), iqr=q3-q1;
  return {lo:q1-1.5*iqr, hi:q3+1.5*iqr};
}
function madMethod(v){
  if(!v.length) return {m:null,mad:null};
  var m=median(v);
  return {m:m, mad:median(v.map(function(x){return Math.abs(x-m);}))};
}
function mannKendall(v){
  var n=v.length;
  if(n<3) return {S:0,Z:null,p:null};
  var S=0;
  for(var i=0;i<n-1;i++) for(var j=i+1;j<n;j++) S+=Math.sign(v[j]-v[i]);
  var varS=(n*(n-1)*(2*n+5))/18;
  var Z=0;
  if(S>0) Z=(S-1)/Math.sqrt(varS); else if(S<0) Z=(S+1)/Math.sqrt(varS);
  var p=2*(1-normCdf(Math.abs(Z)));
  return {S:S,Z:Z,p:p};
}
function normCdf(x){
  var t=1/(1+0.2316419*Math.abs(x));
  var d=0.3989423*Math.exp(-x*x/2);
  var p=d*t*(0.3193815+t*(-0.3565638+t*(1.781478+t*(-1.821256+t*1.330274))));
  if(x>0) p=1-p;
  return p;
}
function entropy(v){
  var s=sum(v); if(s<=0) return null;
  var h=0; for(var i=0;i<v.length;i++) if(v[i]>0){ var p=v[i]/s; h-=p*Math.log(p); } return h;
}
function gini(v){
  var s=v.slice().sort(function(a,b){return a-b;});
  var n=s.length; if(n<2||sum(s)<=0) return null;
  var num=0;
  for(var i=0;i<n;i++) num+=(2*(i+1)-n-1)*s[i];
  return num/(n*sum(s));
}
function crK(v,k){
  var s=v.slice().sort(function(a,b){return b-a;});
  var tot=sum(v); if(tot<=0) return null;
  return sum(s.slice(0,k))/tot*100;
}
function kmeans(pts,k,iters){
  iters=iters||50;
  var n=pts.length;
  if(n<k) return {assign:pts.map(function(_,i){return i;})};
  var centers=[pts[Math.floor(Math.random()*n)].slice()];
  while(centers.length<k){
    var dists=pts.map(function(p){
      var dmin=Infinity;
      for(var c=0;c<centers.length;c++){ var d=0; for(var dim=0;dim<p.length;dim++) d+=(p[dim]-centers[c][dim])*(p[dim]-centers[c][dim]); if(d<dmin) dmin=d; }
      return dmin;
    });
    var total=sum(dists), r=Math.random()*total, acc=0, idx=0;
    for(var i=0;i<n;i++){ acc+=dists[i]; if(acc>=r){ idx=i; break; } }
    centers.push(pts[idx].slice());
  }
  var assign=new Array(n).fill(0);
  for(var it=0;it<iters;it++){
    var changed=false;
    for(var i2=0;i2<n;i2++){
      var best=0,bestD=Infinity;
      for(var c2=0;c2<k;c2++){ var d2=0; for(var dm=0;dm<pts[i2].length;dm++) d2+=(pts[i2][dm]-centers[c2][dm])*(pts[i2][dm]-centers[c2][dm]); if(d2<bestD){bestD=d2;best=c2;} }
      if(assign[i2]!==best){ assign[i2]=best; changed=true; }
    }
    if(!changed) break;
    for(var c3=0;c3<k;c3++){
      var cnt=0, sm=new Array(pts[0].length).fill(0);
      for(var i3=0;i3<n;i3++) if(assign[i3]===c3){ cnt++; for(var d3=0;d3<sm.length;d3++) sm[d3]+=pts[i3][d3]; }
      if(cnt>0) centers[c3]=sm.map(function(x){return x/cnt;});
    }
  }
  return {assign:assign, centers:centers};
}
function pca2d(pts){
  var n=pts.length, d=pts[0].length;
  var means=new Array(d).fill(0);
  for(var i=0;i<n;i++) for(var j=0;j<d;j++) means[j]+=pts[i][j];
  for(var j2=0;j2<d;j2++) means[j2]/=n;
  var centered=pts.map(function(p){ return p.map(function(v,j){ return v-means[j]; }); });
  var C=[];
  for(var a=0;a<d;a++){ C.push([]); for(var b=0;b<d;b++){ var s=0; for(var k=0;k<n;k++) s+=centered[k][a]*centered[k][b]; C[a].push(s/(n-1)); } }
  function powerIter(mat, excl){
    var v=new Array(d).fill(0).map(function(){return Math.random()-0.5;});
    for(var it=0;it<200;it++){
      var w=new Array(d).fill(0);
      for(var i2=0;i2<d;i2++) for(var j3=0;j3<d;j3++) w[i2]+=mat[i2][j3]*v[j3];
      if(excl) for(var e=0;e<excl.length;e++){ var dot=0; for(var q=0;q<d;q++) dot+=w[q]*excl[e][q]; for(var q2=0;q2<d;q2++) w[q2]-=dot*excl[e][q2]; }
      var len=Math.sqrt(sum(w.map(function(x){return x*x;}))); if(len<1e-12) return v;
      v=w.map(function(x){return x/len;});
    }
    return v;
  }
  var v1=powerIter(C,null), v2=powerIter(C,[v1]);
  return {
    pc1:centered.map(function(p){ return sum(p.map(function(x,i){return x*v1[i];})); }),
    pc2:centered.map(function(p){ return sum(p.map(function(x,i){return x*v2[i];})); })
  };
}

/* ---------- СЕРИИ ---------- */
function yearsAll(){
  var m={},res=[];
  for(var i=0;i<state.data.length;i++) m[state.data[i].year]=1;
  for(var k in m) res.push(Number(k));
  res.sort(function(a,b){return a-b;});
  return res;
}
function regionsWithData(){
  var m={},res=[];
  for(var i=0;i<state.data.length;i++) if(!m[state.data[i].region]){ m[state.data[i].region]=1; res.push(state.data[i].region); }
  return res;
}
function seriesFor(r){
  var rows=[];
  for(var i=0;i<state.data.length;i++) if(state.data[i].region===r) rows.push(state.data[i]);
  rows.sort(function(a,b){return a.year-b.year;});
  var ys=[],vs=[];
  for(var j=0;j<rows.length;j++){ ys.push(rows[j].year); vs.push(rows[j].value); }
  return {years:ys,values:vs,rows:rows};
}
function valueAt(r,y){
  for(var i=0;i<state.data.length;i++) if(state.data[i].region===r && state.data[i].year===y) return state.data[i].value;
  return null;
}
function totalsByYearSelected(){
  var map={};
  for(var i=0;i<state.data.length;i++){ var d=state.data[i]; if(state.selected[d.region]) map[d.year]=(map[d.year]||0)+d.value; }
  return map;
}
function totalsByYearAll(){
  var map={};
  for(var i=0;i<state.data.length;i++){ var d=state.data[i]; map[d.year]=(map[d.year]||0)+d.value; }
  return map;
}
function normRegion(s){
  var raw=String(s==null?"":s).trim();
  if(REGIONS.indexOf(raw)>=0) return raw;
  var k=raw.toLowerCase().replace(/\s+/g," ");
  var alias={
    "астрахань":"Астраханская обл.","астраханская область":"Астраханская обл.",
    "волгоград":"Волгоградская обл.","волгоградская область":"Волгоградская обл.",
    "оренбург":"Оренбургская обл.","оренбургская область":"Оренбургская обл.",
    "самара":"Самарская обл.","самарская область":"Самарская обл.",
    "саратов":"Саратовская обл.","саратовская область":"Саратовская обл.",
    "башкирия":"Республика Башкортостан","башкортостан":"Республика Башкортостан",
    "алтай":"Республика Алтай",
    "крым":"Республика Крым",
    "ставрополь":"Ставропольский край","ставропольский край":"Ставропольский край",
    "краснодар":"Краснодарский край","краснодарский край":"Краснодарский край","кубань":"Краснодарский край",
    "ростов":"Ростовская обл.","ростовская область":"Ростовская обл.",
    "белгород":"Белгородская обл.","белгородская область":"Белгородская обл.",
    "дагестан":"Республика Дагестан",
    "чечня":"Чеченская Республика",
    "бурятия":"Республика Бурятия",
    "забайкалье":"Забайкальский край","забайкальский край":"Забайкальский край",
    "приморье":"Приморский край","приморский край":"Приморский край"
  };
  if(alias[k]) return alias[k];
  return raw || null;
}
function ensureRegionInList(r){
  if(REGIONS.indexOf(r)<0){ REGIONS.push(r); renderAllChips(); populateRegionSelects(); }
}

/* ---------- HISTORY ---------- */
var undoStack=[],redoStack=[];
function pushHistory(s){ undoStack.push(JSON.stringify(s)); if(undoStack.length>50) undoStack.shift(); redoStack=[]; updateHistoryUI(); }
function currentSnapshot(){ return { data:state.data, selected:state.selected, theme:state.theme }; }
function applySnapshot(p){
  var o=JSON.parse(p);
  state.data=o.data; state.selected=o.selected;
  refreshAll(); renderAllChips();
}
function undo(){ if(!undoStack.length) return; redoStack.push(JSON.stringify(currentSnapshot())); applySnapshot(undoStack.pop()); updateHistoryUI(); toast("Отменено"); }
function redo(){ if(!redoStack.length) return; pushHistorySnapshot(); applySnapshot(redoStack.pop()); updateHistoryUI(); toast("Повторено"); }
function pushHistorySnapshot(){ undoStack.push(JSON.stringify(currentSnapshot())); updateHistoryUI(); }
function updateHistoryUI(){ if($("historyInfo")) $("historyInfo").textContent=undoStack.length+"/"+(undoStack.length+redoStack.length); }

/* ---------- PERSIST ---------- */
function autosave(){
  if(state.consent!=="all") return;
  try{
    localStorage.setItem("solar_rf_state", JSON.stringify({
      data:state.data, selected:state.selected, theme:state.theme,
      optUnit:state.optUnit, optPrec:state.optPrec, optSep:state.optSep,
      expMode:state.expMode, currentPage:state.currentPage, econ:state.econ
    }));
  }catch(e){}
}
function autoload(){
  if(state.consent!=="all") return false;
  try{
    var raw=localStorage.getItem("solar_rf_state");
    if(!raw) return false;
    var s=JSON.parse(raw);
    if(s.data) state.data=s.data;
    if(s.selected) state.selected=s.selected;
    if(s.theme) state.theme=s.theme;
    if(s.optUnit) state.optUnit=s.optUnit;
    if(s.optPrec) state.optPrec=s.optPrec;
    if(s.optSep) state.optSep=s.optSep;
    if(s.expMode) state.expMode=s.expMode;
    if(s.currentPage) state.currentPage=s.currentPage;
    if(s.econ) state.econ=s.econ;
    return true;
  }catch(e){ return false; }
}

/* ---------- PARSER ---------- */
function parseFlexible(text){
  text=String(text==null?"":text).trim();
  if(!text) throw new Error("Пустая вставка.");
  if(text.charAt(0)==="["||text.charAt(0)==="{"){
    var j=JSON.parse(text);
    if(!Array.isArray(j)){ if(Array.isArray(j.data)) j=j.data; else throw new Error("Ожидался массив."); }
    return j.map(function(o){ return {
      region:o.region||o["Регион"]||"",
      year:Number(o.year||o["Год"]),
      value:Number(o.value||o["Значение"]),
      status:o.status||o["Статус"]||"факт",
      note:o.note||o["Примечание"]||""
    }; });
  }
  var lines=text.split(/\r?\n/).map(function(l){return l.trim();}).filter(Boolean);
  if(!lines.length) throw new Error("Нет строк.");
  var first=lines[0], sep=";";
  if(first.indexOf("\t")>=0) sep="\t";
  else if(first.indexOf(";")>=0) sep=";";
  else if(first.indexOf(",")>=0) sep=",";
  var start=0;
  if(/регион|region|год|year/i.test(first)) start=1;
  var out=[];
  for(var i=start;i<lines.length;i++){
    var c=lines[i].split(sep).map(function(x){return x.trim().replace(/^"|"$/g,"");});
    if(c.length<3) continue;
    out.push({region:c[0],year:Number(c[1]),value:Number(String(c[2]).replace(",",".")),status:c[3]||"факт",note:c[4]||""});
  }
  return out;
}
function validateRows(rows){
  var errors=[],warnings=[],clean=[];
  var seen={}, nowYear=new Date().getFullYear();
  for(var i=0;i<rows.length;i++){
    var r=rows[i], line=i+1;
    if(!r||(!r.region&&(r.year==null||isNaN(r.year))&&(r.value==null||isNaN(r.value)))) continue;
    var reg=normRegion(r.region);
    if(!reg){ errors.push("Строка "+line+": нет региона"); continue; }
    if(!isFinite(r.year)||r.year<1990||r.year>2100){ errors.push("Строка "+line+": год вне 1990…2100"); continue; }
    if(r.year>nowYear+5&&(!r.status||r.status==="факт")) warnings.push("Строка "+line+": будущий год, «факт»");
    if(!isFinite(r.value)||r.value<0){ errors.push("Строка "+line+": значение ≥ 0"); continue; }
    if(r.value>100000){ errors.push("Строка "+line+": значение > 100 000"); continue; }
    var key=reg+"|"+Math.round(r.year);
    if(seen[key]) warnings.push("Строка "+line+": дубликат "+reg+" / "+r.year);
    seen[key]=true;
    var st=String(r.status||"факт").toLowerCase();
    if(STATUSES.indexOf(st)<0) st="факт";
    ensureRegionInList(reg);
    clean.push({region:reg,year:Math.round(r.year),value:Number(r.value),status:st,note:String(r.note||"").trim()});
  }
  return {clean:clean,errors:errors,warnings:warnings};
}
function mergeData(old,add){
  var map={};
  old.forEach(function(r){ map[r.region+"|"+r.year]=r; });
  add.forEach(function(r){ map[r.region+"|"+r.year]=r; });
  var res=[];
  for(var k in map) res.push(map[k]);
  res.sort(function(a,b){ var c=a.region.localeCompare(b.region,"ru"); return c!==0?c:a.year-b.year; });
  return res;
}

/* ---------- UI ---------- */
function renderChips(el){
  if(!el) return;
  el.innerHTML="";
  REGIONS.forEach(function(r){
    var label=document.createElement("label");
    label.className="chip"+(state.selected[r]?" active":"");
    var cb=document.createElement("input");
    cb.type="checkbox"; cb.checked=!!state.selected[r];
    on(cb,"change",function(){
      state.selected[r]=cb.checked;
      label.className="chip"+(cb.checked?" active":"");
      renderAllChips(); refreshAll();
    });
    label.appendChild(cb);
    label.appendChild(document.createTextNode(r));
    el.appendChild(label);
  });
}
function renderAllChips(){ renderChips($("regionChips")); renderChips($("regionChips2")); }
function populateRegionSelects(){
  ["detailRegion","histRegion","passRegion","lagA","lagB","econRegion"].forEach(function(id){
    var s=$(id); if(!s) return;
    var prev=s.value;
    var list=REGIONS.filter(function(r){ return state.selected[r]; });
    if(!list.length) list=regionsWithData();
    s.innerHTML=list.map(function(r){ return '<option>'+escapeHtml(r)+'</option>'; }).join("");
    if(list.indexOf(prev)>=0) s.value=prev;
  });
}
function renderDataTable(){
  var tbody=document.querySelector("#dataTable tbody"); if(!tbody) return;
  if(!state.data.length){ tbody.innerHTML='<tr><td colspan="6" class="empty">Нет данных.</td></tr>'; return; }
  tbody.innerHTML="";
  state.data.forEach(function(row,idx){
    var tr=document.createElement("tr");
    var o=""; REGIONS.forEach(function(r){ o+='<option'+(r===row.region?" selected":"")+'>'+escapeHtml(r)+'</option>'; });
    var so=""; STATUSES.forEach(function(s){ so+='<option'+(s===row.status?" selected":"")+'>'+s+'</option>'; });
    tr.innerHTML=
      '<td><select data-idx="'+idx+'" data-field="region">'+o+'</select></td>'+
      '<td><input type="number" data-idx="'+idx+'" data-field="year" value="'+row.year+'" step="1" min="1990" max="2100"></td>'+
      '<td><input type="number" data-idx="'+idx+'" data-field="value" value="'+row.value+'" step="0.01" min="0"></td>'+
      '<td><select data-idx="'+idx+'" data-field="status">'+so+'</select></td>'+
      '<td><input type="text" data-idx="'+idx+'" data-field="note" value="'+escapeHtml(row.note||"")+'" style="min-width:110px"></td>'+
      '<td><button type="button" class="ghost" data-del="'+idx+'">✕</button></td>';
    tbody.appendChild(tr);
  });
  qsa("#dataTable tbody input,#dataTable tbody select").forEach(function(el){
    on(el,"change",function(ev){
      pushHistory(currentSnapshot());
      var i=+ev.target.getAttribute("data-idx"), f=ev.target.getAttribute("data-field");
      if(f==="year") state.data[i].year=Number(ev.target.value);
      if(f==="value") state.data[i].value=Number(ev.target.value);
      if(f==="region"){ state.data[i].region=ev.target.value; ensureRegionInList(ev.target.value); }
      if(f==="status") state.data[i].status=ev.target.value;
      if(f==="note") state.data[i].note=ev.target.value;
      refreshAll();
    });
  });
  qsa("#dataTable tbody [data-del]").forEach(function(b){
    on(b,"click",function(){ pushHistory(currentSnapshot()); state.data.splice(+b.getAttribute("data-del"),1); refreshAll(); });
  });
}
function renderAppendTable(){
  var tb=document.querySelector("#appendTable tbody"); if(!tb) return;
  if(!state.data.length){ tb.innerHTML='<tr><td colspan="5" class="empty">Пусто</td></tr>'; return; }
  var sorted=state.data.slice().sort(function(a,b){ var c=a.region.localeCompare(b.region,"ru"); return c!==0?c:a.year-b.year; });
  tb.innerHTML=sorted.map(function(d){
    return '<tr><td>'+escapeHtml(d.region)+'</td><td>'+d.year+'</td><td class="num">'+fmtNum(d.value,2)+'</td><td>'+escapeHtml(d.status||"факт")+'</td><td class="wrap">'+escapeHtml(d.note||"")+'</td></tr>';
  }).join("");
}
function buildStats(){
  var rows=[];
  REGIONS.forEach(function(reg){
    if(!state.selected[reg]) return;
    var s=seriesFor(reg); if(!s.values.length) return;
    var totAll=totalsByYearAll(), own=0, all=0;
    for(var i=0;i<s.years.length;i++){ var t=totAll[s.years[i]]; if(t){ own+=s.values[i]; all+=t; } }
    var m=mean(s.values), sd=std(s.values), lr=linreg(s.values);
    rows.push({
      region:reg, n:s.values.length, sum:sum(s.values), mean:m, median:median(s.values), std:sd,
      cv:(sd!=null&&m>0)?sd/m*100:null,
      min:minOf(s.values), max:maxOf(s.values),
      cagr:cagrSeries(s.years,s.values),
      slope:lr?lr.b:null, r2:lr?lr.r2:null,
      share: all>0?own/all*100:null
    });
  });
  return rows;
}
function renderStats(){
  var rows=buildStats();
  var tb=document.querySelector("#statsTable tbody");
  if(tb){
    if(!rows.length) tb.innerHTML='<tr><td colspan="13" class="empty">Нет данных.</td></tr>';
    else tb.innerHTML=rows.map(function(r){
      return '<tr><td>'+escapeHtml(r.region)+'</td><td class="num">'+r.n+'</td>'+
        '<td class="num">'+fmtNum(convertUnits(r.sum),1)+'</td>'+
        '<td class="num">'+fmtNum(convertUnits(r.mean),2)+'</td>'+
        '<td class="num">'+fmtNum(convertUnits(r.median),2)+'</td>'+
        '<td class="num">'+fmtNum(r.std,2)+'</td>'+
        '<td class="num">'+(r.cv!=null?fmtNum(r.cv,2):"—")+'</td>'+
        '<td class="num">'+fmtNum(convertUnits(r.min),2)+'</td>'+
        '<td class="num">'+fmtNum(convertUnits(r.max),2)+'</td>'+
        '<td class="num">'+(r.cagr!=null?fmtNum(r.cagr,2):"—")+'</td>'+
        '<td class="num">'+(r.slope!=null?fmtNum(r.slope,3):"—")+'</td>'+
        '<td class="num">'+(r.r2!=null?fmtNum(r.r2,3):"—")+'</td>'+
        '<td class="num">'+(r.share!=null?fmtNum(r.share,2):"—")+'</td></tr>';
    }).join("");
  }
  var years=yearsAll(), tb2=document.querySelector("#yearTable tbody");
  if(tb2){
    if(!years.length) tb2.innerHTML='<tr><td colspan="6" class="empty">Нет данных.</td></tr>';
    else{
      var tot=totalsByYearSelected(), base=tot[years[0]]||null;
      var arr=years.map(function(y){return tot[y]||0;}), ma=movingAvg(arr,3);
      var prev=null;
      tb2.innerHTML=years.map(function(y,i){
        var v=tot[y]!=null?tot[y]:null;
        var ab=(v!=null&&prev!=null)?v-prev:null;
        var ch=(v!=null&&prev!=null&&prev>0)?(v/prev-1)*100:null;
        var bs=(v!=null&&base!=null&&base>0)?(v/base*100):null;
        var html='<tr><td>'+y+'</td><td class="num">'+(v!=null?fmtNum(convertUnits(v),1):"—")+'</td>'+
          '<td class="num '+(ab>0?'pos':ab<0?'neg':'')+'">'+(ab!=null?fmtNum(convertUnits(ab),2):"—")+'</td>'+
          '<td class="num">'+(ch!=null?fmtNum(ch,2):"—")+'</td>'+
          '<td class="num">'+(bs!=null?fmtNum(bs,2):"—")+'</td>'+
          '<td class="num">'+fmtNum(ma[i],2)+'</td></tr>';
        prev=v; return html;
      }).join("");
    }
  }
  var sel=$("detailRegion");
  if(sel){
    var pv=sel.value;
    sel.innerHTML=rows.map(function(r){return '<option>'+escapeHtml(r.region)+'</option>';}).join("");
    if(rows.some(function(r){return r.region===pv;})) sel.value=pv;
    renderDetail();
  }
}
function renderDetail(){
  var reg=$("detailRegion")?$("detailRegion").value:"";
  var tb=document.querySelector("#detailTable tbody"); if(!tb) return;
  if(!reg){ tb.innerHTML='<tr><td colspan="7" class="empty">Выберите регион.</td></tr>'; return; }
  var s=seriesFor(reg);
  if(!s.values.length){ tb.innerHTML='<tr><td colspan="7" class="empty">Нет данных.</td></tr>'; return; }
  var m=mean(s.values), sd=std(s.values), base=s.values[0];
  var prev=null;
  tb.innerHTML=s.years.map(function(y,i){
    var v=s.values[i];
    var ab=prev!=null?v-prev:null;
    var ch=(prev!=null&&prev>0)?(v/prev-1)*100:null;
    var bs=base>0?v/base*100:null;
    var z=(sd!=null&&sd>0)?(v-m)/sd:null;
    var fl=z==null?"":(Math.abs(z)>3?"сильный":(Math.abs(z)>2?"выброс":""));
    var html='<tr><td>'+y+'</td><td class="num">'+fmtNum(convertUnits(v),2)+'</td>'+
      '<td class="num '+(ab>0?'pos':ab<0?'neg':'')+'">'+(ab!=null?fmtNum(convertUnits(ab),2):"—")+'</td>'+
      '<td class="num">'+(ch!=null?fmtNum(ch,2):"—")+'</td>'+
      '<td class="num">'+(bs!=null?fmtNum(bs,2):"—")+'</td>'+
      '<td class="num '+(fl?'warn':'')+'">'+(z!=null?fmtNum(z,2):"—")+'</td>'+
      '<td>'+fl+'</td></tr>';
    prev=v; return html;
  }).join("");
}

/* ---------- PLOT ---------- */
function gridColor(){ return state.theme==="dark"?"#1a2637":"#e1e7f2"; }
var PLOT_CFG = {responsive:true, displayModeBar:false};
function baseLayout(extra){
  var dark=state.theme==="dark";
  var base={
    margin:{l:60,r:20,t:24,b:60},
    paper_bgcolor: dark?"#0d1726":"#ffffff",
    plot_bgcolor:  dark?"#0d1726":"#ffffff",
    font:{family:'"Inter",Arial,sans-serif', color: dark?"#e8eef8":"#0a1628", size:11},
    xaxis:{title:{text:"Год"}, gridcolor:gridColor(), zeroline:false},
    yaxis:{title:{text:unitLabel()}, gridcolor:gridColor(), zeroline:false},
    legend:{orientation:"h", y:-0.18, font:{size:10}},
    hovermode:"x unified",
    dragmode:"zoom"
  };
  if(extra) for(var k in extra) base[k]=extra[k];
  return base;
}
function renderMainChart(){
  var div=$("chart"); if(!div) return;
  if(!plotlyReady()){ div.innerHTML='<div class="empty">Plotly не загружен.</div>'; return; }
  if(!state.data.length){ Plotly.newPlot(div,[],baseLayout({title:{text:"Нет данных"}}),PLOT_CFG); return; }
  var sel=REGIONS.filter(function(r){return state.selected[r];});
  if(!sel.length){ Plotly.newPlot(div,[],baseLayout({title:{text:"Выберите регионы"}}),PLOT_CFG); return; }
  var type=$("chartType").value, yScale=$("yScale").value;
  var yMin=Number($("yearMin").value)||null, yMax=Number($("yearMax").value)||null;
  var years=yearsAll().filter(function(y){return (yMin==null||y>=yMin)&&(yMax==null||y<=yMax);});
  if(type==="stack100"){
    var totals={};
    years.forEach(function(y){ var t=0; sel.forEach(function(r){ var v=valueAt(r,y); if(v!=null) t+=v; }); totals[y]=t; });
    var tr=sel.map(function(r,i){
      var ys=years.map(function(y){ var v=valueAt(r,y), t=totals[y]; return (v!=null&&t>0)?v/t*100:null; });
      return {name:r, x:years, y:ys, type:"bar", marker:{color:PALETTE[i%PALETTE.length]}};
    });
    Plotly.newPlot(div,tr,baseLayout({barmode:"stack", yaxis:{title:{text:"Доля, %"}, gridcolor:gridColor(), ticksuffix:"%", range:[0,100]}}),PLOT_CFG);
    return;
  }
  var tr2=sel.map(function(r,i){
    var ys=years.map(function(y){ var v=valueAt(r,y); return v!=null?convertUnits(v):null; });
    var color=PALETTE[i%PALETTE.length];
    if(type==="lines") return {name:r, x:years, y:ys, type:"scatter", mode:"lines+markers", line:{width:2.2,color:color}, marker:{size:5,color:color}};
    if(type==="area") return {name:r, x:years, y:ys, type:"scatter", mode:"lines+markers", fill:"tozeroy", line:{width:1.8,color:color}, marker:{size:5,color:color}};
    return {name:r, x:years, y:ys, type:"bar", marker:{color:color}};
  });
  Plotly.newPlot(div,tr2,baseLayout({
    barmode: type==="bar"?"group":undefined,
    xaxis:{title:{text:"Год"}, gridcolor:gridColor(), zeroline:false, tickmode:"array", tickvals:years},
    yaxis:{title:{text:unitLabel()}, gridcolor:gridColor(), zeroline:false, type:yScale}
  }),PLOT_CFG);
}
function renderCompare(){
  var div=$("compareChart"); if(!div||!plotlyReady()) return;
  var sel=REGIONS.filter(function(r){return state.selected[r];});
  var tb=document.querySelector("#compareTable tbody");
  if(!state.data.length||!sel.length){
    if(tb) tb.innerHTML='<tr><td colspan="5" class="empty">Нет данных.</td></tr>';
    Plotly.newPlot(div,[],baseLayout({title:{text:"Нет данных"}}),PLOT_CFG); return;
  }
  var traces=[], rows=[];
  sel.forEach(function(r,i){
    var s=seriesFor(r); if(!s.values.length) return;
    var base=s.values[0]; if(base<=0) return;
    var idx=s.values.map(function(v){return v/base*100;});
    var color=PALETTE[i%PALETTE.length];
    traces.push({name:r, x:s.years, y:idx, type:"scatter", mode:"lines+markers", line:{width:2.2,color:color}, marker:{size:4,color:color}});
    rows.push({region:r, baseYear:s.years[0], last:idx[idx.length-1], delta:idx[idx.length-1]-100});
  });
  Plotly.newPlot(div,traces,baseLayout({yaxis:{title:{text:"Индекс"},gridcolor:gridColor(),zeroline:false}}),PLOT_CFG);
  if(tb) tb.innerHTML=rows.map(function(r){
    return '<tr><td>'+escapeHtml(r.region)+'</td><td class="num">'+r.baseYear+'</td><td class="num">100,0</td><td class="num">'+fmtNum(r.last,1)+'</td><td class="num '+(r.delta>0?'pos':r.delta<0?'neg':'')+'">'+fmtNum(r.delta,1)+'</td></tr>';
  }).join("");
}
function pickBestModel(y){
  var lr=linreg(y), poly=polyreg(y,2), exp=expreg(y);
  var models=[];
  if(lr){ var fL=y.map(function(_,i){return lr.a+lr.b*i;}); models.push({name:"Линейный",r2:lr.r2,mape:mape(y,fL),a:lr.a,b:lr.b,predict:function(t){return lr.a+lr.b*t;}}); }
  if(poly){ var fP=y.map(function(_,i){return poly.predict(i);}); models.push({name:"Полином-2",r2:poly.r2,mape:mape(y,fP),a:poly.coef[0],b:poly.coef[1],predict:poly.predict}); }
  if(exp){ var fE=y.map(function(_,i){return exp.predict(i);}); models.push({name:"Экспонента",r2:exp.r2,mape:mape(y,fE),a:exp.a,b:exp.b,predict:exp.predict}); }
  if(!models.length) return null;
  models.sort(function(a,b){return (b.r2||0)-(a.r2||0);});
  return models[0];
}
function renderTrend(){
  var rows=buildStats().filter(function(r){return r.slope!=null;});
  var tb=document.querySelector("#trendTable tbody");
  var div=$("trendChart");
  if(!rows.length){ if(tb) tb.innerHTML='<tr><td colspan="10" class="empty">Недостаточно данных.</td></tr>'; if(div&&plotlyReady()) Plotly.newPlot(div,[],baseLayout({title:{text:"Нет данных"}}),PLOT_CFG); return; }
  var horizon=Number($("forecastYears").value)||3;
  if(div&&plotlyReady()){
    var traces=[];
    rows.forEach(function(r,i){
      var s=seriesFor(r.region); var color=PALETTE[i%PALETTE.length];
      traces.push({name:r.region, x:s.years, y:s.values.map(convertUnits), type:"scatter", mode:"lines+markers", line:{color:color,width:2}, marker:{color:color,size:4}});
      var lr=linreg(s.values);
      if(lr){
        var lastY=s.years[s.years.length-1], fx=[], fy=[];
        for(var k=0;k<s.values.length;k++){ fx.push(s.years[k]); fy.push(lr.a+lr.b*k); }
        for(var k2=1;k2<=horizon;k2++){ fx.push(lastY+k2); fy.push(lr.a+lr.b*(s.values.length-1+k2)); }
        traces.push({name:r.region+" тренд", x:fx, y:fy.map(convertUnits), type:"scatter", mode:"lines", line:{color:color,width:1.2,dash:"dot"}, showlegend:false, hoverinfo:"skip"});
      }
    });
    Plotly.newPlot(div,traces,baseLayout({}),PLOT_CFG);
  }
  if(tb) tb.innerHTML=rows.map(function(r){
    var s=seriesFor(r.region), best=pickBestModel(s.values); if(!best) return '';
    var pred=function(k){return best.predict(s.values.length-1+k);};
    return '<tr><td>'+escapeHtml(r.region)+'</td><td>'+best.name+'</td>'+
      '<td class="num">'+fmtNum(best.r2,3)+'</td><td class="num">'+(best.mape!=null?fmtNum(best.mape,2):"—")+'</td>'+
      '<td class="num">'+fmtNum(best.a,3)+'</td><td class="num">'+fmtNum(best.b,4)+'</td>'+
      '<td class="num">'+fmtNum(convertUnits(pred(1)),2)+'</td>'+
      '<td class="num">'+fmtNum(convertUnits(pred(2)),2)+'</td>'+
      '<td class="num">'+fmtNum(convertUnits(pred(3)),2)+'</td>'+
      '<td class="num">'+fmtNum(convertUnits(pred(5)),2)+'</td></tr>';
  }).join("");
}
function renderStructure(){
  var sel=REGIONS.filter(function(r){return state.selected[r];});
  var years=yearsAll(), tb=document.querySelector("#structureTable tbody"), div=$("structureChart");
  if(!state.data.length||!sel.length){ if(tb) tb.innerHTML='<tr><td colspan="4" class="empty">Нет данных.</td></tr>'; if(div&&plotlyReady()) Plotly.newPlot(div,[],baseLayout({title:{text:"Нет данных"}}),PLOT_CFG); return; }
  var totals=totalsByYearAll();
  if(plotlyReady()&&div){
    var traces=sel.map(function(r,i){
      var ys=years.map(function(y){ var v=valueAt(r,y), t=totals[y]; return (v!=null&&t>0)?v/t*100:null; });
      return {name:r, x:years, y:ys, type:"bar", marker:{color:PALETTE[i%PALETTE.length]}};
    });
    Plotly.newPlot(div,traces,baseLayout({barmode:"stack", yaxis:{title:{text:"Доля, %"}, gridcolor:gridColor(), ticksuffix:"%", range:[0,100]}}),PLOT_CFG);
  }
  if(tb) tb.innerHTML=years.map(function(y){
    var shares=sel.map(function(r){ var v=valueAt(r,y),t=totals[y]; return (v!=null&&t>0)?v/t*100:null; }).filter(function(x){return x!=null;});
    var tot=sum(shares);
    var hhi=tot>0?sum(shares.map(function(s){return Math.pow(s*100/tot,2);})):null;
    var interp=hhi==null?"—":(hhi<1000?"низкая":hhi<1800?"умеренная":hhi<2500?"высокая":"очень высокая");
    var list=sel.map(function(r){ var v=valueAt(r,y),t=totals[y]; return (v!=null&&t>0)?escapeHtml(r)+": "+fmtNum(v/t*100,1):null; }).filter(Boolean).join(" · ");
    return '<tr><td>'+y+'</td><td class="num">'+(hhi!=null?fmtNum(hhi,0):"—")+'</td><td>'+interp+'</td><td class="wrap">'+(list||"—")+'</td></tr>';
  }).join("");
}
function renderAnomaly(){
  var sel=REGIONS.filter(function(r){return state.selected[r];});
  var tb=document.querySelector("#anomalyTable tbody"); var rows=[];
  sel.forEach(function(reg){
    var s=seriesFor(reg); if(s.values.length<2) return;
    var m=mean(s.values), sd=std(s.values), iqr=iqrMethod(s.values), mad=madMethod(s.values);
    for(var i=0;i<s.values.length;i++){
      var v=s.values[i];
      var z=sd?(v-m)/sd:null;
      var iq=(iqr.lo!=null&&(v<iqr.lo||v>iqr.hi));
      var mf=(mad.mad&&Math.abs(v-mad.m)>3*mad.mad);
      var flags=[];
      if(z!=null&&Math.abs(z)>2) flags.push("z"+(Math.abs(z)>3?"!!":"!"));
      if(iq) flags.push("IQR");
      if(mf) flags.push("MAD");
      if(flags.length) rows.push({region:reg,year:s.years[i],value:v,z:z,iqrFlag:iq,madFlag:mf,flags:flags.join(" · ")});
    }
  });
  if(tb){
    if(!rows.length) tb.innerHTML='<tr><td colspan="7" class="empty">Аномалий нет.</td></tr>';
    else{
      rows.sort(function(a,b){return Math.abs(b.z||0)-Math.abs(a.z||0);});
      tb.innerHTML=rows.map(function(r){
        return '<tr><td>'+escapeHtml(r.region)+'</td><td>'+r.year+'</td><td class="num">'+fmtNum(convertUnits(r.value),2)+'</td>'+
          '<td class="num warn">'+(r.z!=null?fmtNum(r.z,2):"—")+'</td>'+
          '<td>'+(r.iqrFlag?'✓':'—')+'</td><td>'+(r.madFlag?'✓':'—')+'</td><td class="warn">'+r.flags+'</td></tr>';
      }).join("");
    }
  }
  var wrap=$("corrTableWrap"); if(!wrap) return;
  if(sel.length<2){ wrap.innerHTML='<div class="empty">≥ 2 регионов.</div>'; return; }
  var years=yearsAll().filter(function(y){ for(var i=0;i<sel.length;i++) if(valueAt(sel[i],y)==null) return false; return true; });
  if(years.length<2){ wrap.innerHTML='<div class="empty">Мало общих лет.</div>'; return; }
  var map={}; sel.forEach(function(r){ map[r]=years.map(function(y){return valueAt(r,y);}); });
  var html='<div class="tbl-wrap small"><table><thead><tr><th></th>';
  sel.forEach(function(r){ html+='<th>'+escapeHtml(r)+'</th>'; });
  html+='</tr></thead><tbody>';
  sel.forEach(function(r1){
    html+='<tr><th>'+escapeHtml(r1)+'</th>';
    sel.forEach(function(r2){
      var c=r1===r2?1:pearson(map[r1],map[r2]);
      var cls=c==null?"":(c>0.7?"pos":(c<-0.7?"neg":""));
      html+='<td class="num '+cls+'">'+(c!=null?fmtNum(c,3):"—")+'</td>';
    });
    html+='</tr>';
  });
  html+='</tbody></table></div>';
  wrap.innerHTML=html;
}
function renderHeatmap(){
  var sel=REGIONS.filter(function(r){return state.selected[r];});
  var years=yearsAll(), div=$("heatmapChart");
  if(div&&plotlyReady()){
    if(sel.length<2) Plotly.newPlot(div,[],baseLayout({title:{text:"≥ 2 регионов"}}),PLOT_CFG);
    else{
      var common=years.filter(function(y){ for(var i=0;i<sel.length;i++) if(valueAt(sel[i],y)==null) return false; return true; });
      if(common.length<2) Plotly.newPlot(div,[],baseLayout({title:{text:"Мало общих лет"}}),PLOT_CFG);
      else{
        var matrix=sel.map(function(r1){ return sel.map(function(r2){ var a=common.map(function(y){return valueAt(r1,y);}), b=common.map(function(y){return valueAt(r2,y);}); var c=(r1===r2)?1:pearson(a,b); return c==null?0:c; }); });
        Plotly.newPlot(div,[{type:"heatmap", z:matrix, x:sel, y:sel, colorscale:[[0,"#dc2626"],[0.5,"#f3f4f6"],[1,"#1e40af"]], zmin:-1, zmax:1, colorbar:{title:"r"}}],
          baseLayout({xaxis:{title:{text:""}, gridcolor:gridColor(), tickangle:-30}, yaxis:{title:{text:""}, gridcolor:gridColor(), autorange:"reversed"}}),PLOT_CFG);
      }
    }
  }
  var md=$("matrixChart");
  if(md&&plotlyReady()&&sel.length){
    var z=sel.map(function(r){ return years.map(function(y){ var v=valueAt(r,y); return v!=null?convertUnits(v):null; }); });
    Plotly.newPlot(md,[{type:"heatmap", z:z, x:years, y:sel, colorscale:"Viridis", colorbar:{title:unitLabel()}}],
      baseLayout({xaxis:{title:{text:"Год"}, gridcolor:gridColor()}, yaxis:{title:{text:""}, gridcolor:gridColor(), autorange:"reversed"}}),PLOT_CFG);
  }
  var hs=$("histRegion");
  if(hs&&plotlyReady()){
    var pv=hs.value;
    hs.innerHTML=sel.map(function(r){return '<option>'+escapeHtml(r)+'</option>';}).join("");
    if(sel.indexOf(pv)>=0) hs.value=pv;
    var reg=hs.value, s=seriesFor(reg);
    if($("histChart")){
      Plotly.newPlot($("histChart"),[{type:"histogram", x:s.values.map(convertUnits), marker:{color:PALETTE[0]}, nbinsx:Math.max(3,Math.min(12,Math.round(Math.sqrt(s.values.length))))}],
        baseLayout({xaxis:{title:{text:unitLabel()}, gridcolor:gridColor()}, yaxis:{title:{text:"Частота"}, gridcolor:gridColor()}, title:{text:"Распределение: "+reg}}),PLOT_CFG);
    }
  }
  var wd=$("waterfallChart");
  if(wd&&plotlyReady()){
    var tot=totalsByYearSelected();
    var steps=[];
    for(var i=1;i<years.length;i++){ steps.push({y:years[i], diff:(tot[years[i]]||0)-(tot[years[i-1]]||0)}); }
    if(steps.length){
      var xs=steps.map(function(s){return String(s.y);}), ys=steps.map(function(s){return convertUnits(s.diff);});
      var colors=ys.map(function(v){return v>=0?"#059669":"#dc2626";});
      Plotly.newPlot(wd,[{x:xs, y:ys, type:"bar", marker:{color:colors}}],
        baseLayout({xaxis:{title:{text:"Год"}, gridcolor:gridColor()}, yaxis:{title:{text:"Δ "+unitLabel()}, gridcolor:gridColor()}, title:{text:"Вклад года в прирост"}}),PLOT_CFG);
    } else Plotly.newPlot(wd,[],baseLayout({title:{text:"Мало данных"}}),PLOT_CFG);
  }
  var bd=$("boxChart");
  if(bd&&plotlyReady()&&sel.length){
    var tr=sel.map(function(r,i){
      var s=seriesFor(r);
      return {y:s.values.map(convertUnits), name:r, type:"box", marker:{color:PALETTE[i%PALETTE.length]}};
    });
    Plotly.newPlot(bd,tr,baseLayout({yaxis:{title:{text:unitLabel()}, gridcolor:gridColor()}, xaxis:{title:{text:""}, gridcolor:gridColor(), tickangle:-20}}),PLOT_CFG);
  }
}
function renderAdvanced(){
  var sel=REGIONS.filter(function(r){return state.selected[r];});
  var years=yearsAll();
  if(!state.data.length||!sel.length){
    ["advGini","advCR3","advCR5","advEntropy","advClusters","advPCA"].forEach(function(id){ if($(id)) $(id).textContent="—"; });
    return;
  }
  var lastY=years[years.length-1];
  var vals=sel.map(function(r){return valueAt(r,lastY);}).filter(function(v){return v!=null;});
  var g=gini(vals), cr3=crK(vals,3), cr5=crK(vals,5), h=entropy(vals);
  $("advGini").textContent=g!=null?fmtNum(g,3):"—";
  $("advCR3").textContent=cr3!=null?fmtNum(cr3,1)+"%":"—";
  $("advCR5").textContent=cr5!=null?fmtNum(cr5,1)+"%":"—";
  $("advEntropy").textContent=h!=null?fmtNum(h,3):"—";
  var vecs=[];
  sel.forEach(function(r){
    var s=seriesFor(r);
    if(s.values.length>=3){ var mx=maxOf(s.values)||1; vecs.push({region:r, vec:s.values.map(function(v){return v/mx;})}); }
  });
  if(vecs.length<3){ $("advClusters").textContent="—"; $("advPCA").textContent="—"; return; }
  var minLen=Math.min.apply(null,vecs.map(function(v){return v.vec.length;}));
  var pts=vecs.map(function(v){return v.vec.slice(0,minLen);});
  var pca=pca2d(pts), km=kmeans(pts,Math.min(3,vecs.length));
  $("advClusters").textContent=Math.min(3,vecs.length);
  $("advPCA").textContent="PC1/PC2";
  var colors=["#f59e0b","#dc2626","#059669","#1e40af"];
  if($("pcaChart")&&plotlyReady()){
    var traces=[];
    for(var c=0;c<Math.min(3,vecs.length);c++){
      var idx=[];
      for(var i=0;i<km.assign.length;i++) if(km.assign[i]===c) idx.push(i);
      traces.push({x:idx.map(function(i){return pca.pc1[i];}), y:idx.map(function(i){return pca.pc2[i];}), text:idx.map(function(i){return vecs[i].region;}), mode:"markers+text", textposition:"top center", type:"scatter", name:"Кластер "+(c+1), marker:{size:12, color:colors[c]}});
    }
    Plotly.newPlot($("pcaChart"),traces,baseLayout({xaxis:{title:{text:"PC1"}, gridcolor:gridColor()}, yaxis:{title:{text:"PC2"}, gridcolor:gridColor()}, title:{text:"PCA + k-means"}}),PLOT_CFG);
  }
  var ct=document.querySelector("#clusterTable tbody");
  if(ct) ct.innerHTML=vecs.map(function(v,i){
    return '<tr><td>'+escapeHtml(v.region)+'</td><td>'+(km.assign[i]+1)+'</td><td class="num">'+fmtNum(pca.pc1[i],3)+'</td><td class="num">'+fmtNum(pca.pc2[i],3)+'</td></tr>';
  }).join("");
  var rd=$("radarChart");
  if(rd&&plotlyReady()){
    var rows=buildStats();
    if(rows.length){
      var maxMean=Math.max.apply(null,rows.map(function(r){return r.mean||0;}));
      var maxCv=Math.max.apply(null,rows.map(function(r){return Math.abs(r.cv||0);}));
      var maxCagr=Math.max.apply(null,rows.map(function(r){return Math.abs(r.cagr||0);}));
      var maxShare=Math.max.apply(null,rows.map(function(r){return r.share||0;}));
      var tr=rows.slice(0,6).map(function(r,i){
        return {type:"scatterpolar", r:[(r.mean||0)/(maxMean||1), 1-Math.min(1,Math.abs(r.cv||0)/(maxCv||1)), Math.abs(r.cagr||0)/(maxCagr||1), (r.share||0)/(maxShare||1)], theta:["Среднее","CV","CAGR","Доля"], fill:"toself", name:r.region, line:{color:PALETTE[i%PALETTE.length]}};
      });
      Plotly.newPlot(rd,tr,{polar:{radialaxis:{visible:true, range:[0,1]}}, paper_bgcolor:state.theme==="dark"?"#0d1726":"#fff", plot_bgcolor:state.theme==="dark"?"#0d1726":"#fff", font:{color:state.theme==="dark"?"#e8eef8":"#0a1628"}, showlegend:true},PLOT_CFG);
    }
  }
  var mkt=document.querySelector("#mkTable tbody");
  if(mkt) mkt.innerHTML=sel.map(function(reg){
    var s=seriesFor(reg), mk=mannKendall(s.values);
    var tr=mk.Z==null?"—":(mk.Z>1.96?"рост":(mk.Z<-1.96?"падение":"нет"));
    return '<tr><td>'+escapeHtml(reg)+'</td><td class="num">'+mk.S+'</td><td class="num">'+(mk.Z!=null?fmtNum(mk.Z,2):"—")+'</td><td class="num">'+(mk.p!=null?fmtNum(mk.p,4):"—")+'</td><td>'+tr+'</td></tr>';
  }).join("");
  var acd=$("acfChart");
  if(acd&&plotlyReady()){
    var reg=sel[0], s=seriesFor(reg);
    if(s.values.length>=4){
      var m=mean(s.values), v=s.values.map(function(x){return x-m;}), den=sum(v.map(function(x){return x*x;})), a=[];
      for(var k=0;k<Math.min(10,v.length-1);k++){
        var num=0; for(var i2=0;i2<v.length-k;i2++) num+=v[i2]*v[i2+k];
        a.push(num/(den||1));
      }
      Plotly.newPlot(acd,[{x:a.map(function(_,i){return i;}), y:a, type:"bar", marker:{color:PALETTE[0]}}],
        baseLayout({xaxis:{title:{text:"Лаг"}, gridcolor:gridColor()}, yaxis:{title:{text:"ACF"}, gridcolor:gridColor()}, title:{text:"ACF: "+reg}}),PLOT_CFG);
    } else Plotly.newPlot(acd,[],baseLayout({title:{text:"Мало данных"}}),PLOT_CFG);
  }
  var la=$("lagA"), lb=$("lagB");
  if(la&&lb&&!la.options.length){
    la.innerHTML=sel.map(function(r){return '<option>'+escapeHtml(r)+'</option>';}).join("");
    lb.innerHTML=sel.map(function(r){return '<option>'+escapeHtml(r)+'</option>';}).join("");
    if(sel.length>1) lb.value=sel[1];
  }
}
function buildLagChart(){
  var a=$("lagA").value, b=$("lagB").value; if(!a||!b) return;
  var years=yearsAll();
  var common=years.filter(function(y){return valueAt(a,y)!=null&&valueAt(b,y)!=null;});
  if(common.length<4){ if(plotlyReady()) Plotly.newPlot($("lagChart"),[],baseLayout({title:{text:"Мало общих лет"}}),PLOT_CFG); return; }
  var res=[0,1,2,3].map(function(lag){
    var x=[],y=[];
    for(var i=0;i<common.length-lag;i++){ x.push(valueAt(a,common[i])); y.push(valueAt(b,common[i+lag])); }
    return {lag:lag, r:pearson(x,y)};
  });
  if(plotlyReady()){
    Plotly.newPlot($("lagChart"),[{x:res.map(function(r){return r.lag;}), y:res.map(function(r){return r.r;}), type:"bar", marker:{color:PALETTE[0]}}],
      baseLayout({xaxis:{title:{text:"Лаг"}, gridcolor:gridColor()}, yaxis:{title:{text:"r"}, gridcolor:gridColor(), range:[-1,1]}, title:{text:"Корреляция "+a+" → "+b}}),PLOT_CFG);
  }
}
function renderHealth(){
  if(!$("healthGrid")) return;
  var years=yearsAll(), regions=regionsWithData(), issues=[], missing=0;
  regions.forEach(function(r){
    var s=seriesFor(r), set={}; s.years.forEach(function(y){set[y]=1;});
    var gaps=years.filter(function(y){return !set[y];});
    if(gaps.length){ missing+=gaps.length; issues.push({region:r,type:"Пропуски лет",detail:gaps.join(", ")}); }
    if(s.values.length<2) issues.push({region:r,type:"Короткий ряд",detail:s.values.length+" лет"});
    for(var i=1;i<s.values.length;i++) if(s.values[i-1]>0){
      var ch=(s.values[i]-s.values[i-1])/s.values[i-1];
      if(Math.abs(ch)>0.5) issues.push({region:r,type:"Резкий скачок",detail:s.years[i-1]+" → "+s.years[i]+": "+fmtNum(ch*100,0)+"%"});
    }
    s.rows.forEach(function(row){ if(row.status==="факт"&&row.year>new Date().getFullYear()) issues.push({region:r,type:"Факт в будущем",detail:row.year}); });
  });
  var empty=REGIONS.filter(function(r){return regions.indexOf(r)<0;});
  if(empty.length) issues.push({region:"—",type:"Регионы без данных",detail:empty.join(", ")});
  var total=REGIONS.length*years.length, fill=total?state.data.length/total*100:0;
  $("healthGrid").innerHTML=
    '<div class="health-item '+(state.data.length?'ok':'warn')+'"><b>'+state.data.length+'</b>строк</div>'+
    '<div class="health-item '+(regions.length===REGIONS.length?'ok':'warn')+'"><b>'+regions.length+' / '+REGIONS.length+'</b>регионов</div>'+
    '<div class="health-item '+(fill>80?'ok':fill>50?'warn':'err')+'"><b>'+fmtNum(fill,1)+'%</b>заполнения</div>'+
    '<div class="health-item '+(missing?'warn':'ok')+'"><b>'+missing+'</b>пропусков</div>'+
    '<div class="health-item '+(issues.length?'warn':'ok')+'"><b>'+issues.length+'</b>проблем</div>';
  var html='<table><thead><tr><th>Регион</th>';
  years.forEach(function(y){ html+='<th class="num">'+y+'</th>'; });
  html+='</tr></thead><tbody>';
  REGIONS.forEach(function(r){
    html+='<tr><td>'+escapeHtml(r)+'</td>';
    years.forEach(function(y){
      var v=valueAt(r,y);
      html+='<td>'+(v!=null?'<span class="matrix-cell has" title="'+r+' '+y+'"></span>':'<span class="matrix-cell no"></span>')+'</td>';
    });
    html+='</tr>';
  });
  html+='</tbody></table>';
  $("healthMatrix").innerHTML=html;
  if(!issues.length) $("healthIssues").innerHTML='<div class="empty">Проблем не обнаружено.</div>';
  else{
    var ih='<table><thead><tr><th>Регион</th><th>Тип</th><th>Детали</th></tr></thead><tbody>';
    issues.forEach(function(it){ ih+='<tr><td>'+escapeHtml(it.region)+'</td><td>'+escapeHtml(it.type)+'</td><td class="wrap">'+escapeHtml(it.detail)+'</td></tr>'; });
    ih+='</tbody></table>';
    $("healthIssues").innerHTML=ih;
  }
}
function renderRegionPassport(){
  var sel=$("passRegion"); if(!sel) return;
  var list=REGIONS.filter(function(r){return state.selected[r];}); if(!list.length) list=regionsWithData();
  sel.innerHTML=list.map(function(r){return '<option>'+escapeHtml(r)+'</option>';}).join("");
  var reg=sel.value; if(!reg) return;
  var s=seriesFor(reg), pk=$("passKpis");
  if(!s.values.length){ pk.innerHTML='<div class="empty">Нет данных</div>'; return; }
  var m=mean(s.values), sd=std(s.values), cg=cagrSeries(s.years,s.values);
  pk.innerHTML=
    '<div class="kpi"><div class="label">Строк</div><div class="val">'+s.values.length+'</div></div>'+
    '<div class="kpi"><div class="label">Σ</div><div class="val">'+fmtNum(convertUnits(sum(s.values)),1)+'</div></div>'+
    '<div class="kpi"><div class="label">Среднее</div><div class="val">'+fmtNum(convertUnits(m),2)+'</div></div>'+
    '<div class="kpi"><div class="label">σ</div><div class="val">'+fmtNum(sd,2)+'</div></div>'+
    '<div class="kpi"><div class="label">CAGR</div><div class="val">'+(cg!=null?fmtNum(cg,2)+"%":"—")+'</div></div>';
  if(plotlyReady()) Plotly.newPlot($("passChart"),[{name:reg, x:s.years, y:s.values.map(convertUnits), type:"scatter", mode:"lines+markers", line:{width:2.2,color:PALETTE[0]}, marker:{size:5,color:PALETTE[0]}}],
    baseLayout({title:{text:"Паспорт: "+reg}}),PLOT_CFG);
  var tb=document.querySelector("#passTable tbody");
  if(tb){
    var prev=null;
    tb.innerHTML=s.years.map(function(y,i){
      var v=s.values[i], ab=prev!=null?v-prev:null, ch=(prev!=null&&prev>0)?(v/prev-1)*100:null;
      var bs=s.values[0]>0?v/s.values[0]*100:null, z=sd?(v-m)/sd:null;
      var row=s.rows[i];
      var html='<tr><td>'+y+'</td><td class="num">'+fmtNum(convertUnits(v),2)+'</td>'+
        '<td class="num '+(ab>0?'pos':ab<0?'neg':'')+'">'+(ab!=null?fmtNum(convertUnits(ab),2):"—")+'</td>'+
        '<td class="num">'+(ch!=null?fmtNum(ch,2):"—")+'</td>'+
        '<td class="num">'+(bs!=null?fmtNum(bs,2):"—")+'</td>'+
        '<td class="num '+(z!=null&&Math.abs(z)>2?'warn':'')+'">'+(z!=null?fmtNum(z,2):"—")+'</td>'+
        '<td>'+escapeHtml(row.status||"факт")+'</td><td class="wrap">'+escapeHtml(row.note||"")+'</td></tr>';
      prev=v; return html;
    }).join("");
  }
}
function renderScenario(){
  var rate=Number($("scenRate").value)/100, years=Number($("scenYears").value);
  var yl=yearsAll();
  if(!yl.length){ $("scenKpis").innerHTML='<div class="empty">Нет данных</div>'; return; }
  var lastY=yl[yl.length-1], tot=totalsByYearSelected(), base=tot[lastY]||0;
  var arr=[{year:lastY,val:base}];
  for(var i=1;i<=years;i++) arr.push({year:lastY+i, val:base*Math.pow(1+rate,i)});
  var last=arr[arr.length-1].val;
  $("scenKpis").innerHTML=
    '<div class="kpi"><div class="label">База '+lastY+'</div><div class="val">'+fmtNum(convertUnits(base),1)+'</div></div>'+
    '<div class="kpi"><div class="label">Прогноз '+(lastY+years)+'</div><div class="val">'+fmtNum(convertUnits(last),1)+'</div></div>'+
    '<div class="kpi"><div class="label">Прирост</div><div class="val '+(last>=base?'pos':'neg')+'">'+(base>0?fmtNum((last/base-1)*100,1)+"%":"—")+'</div></div>';
  if(plotlyReady()) Plotly.newPlot($("scenChart"),[{x:arr.map(function(a){return a.year;}), y:arr.map(function(a){return convertUnits(a.val);}), type:"scatter", mode:"lines+markers", line:{width:2.2,color:PALETTE[0]}, marker:{size:6,color:PALETTE[0]}}],
    baseLayout({title:{text:"Сценарий: "+(rate*100).toFixed(1)+"%/год"}}),PLOT_CFG);
}

/* ============================================================
   ЭКОНОМИКА СЭС — v6.1 (полная переработка)
   Методология: IEA LCOE, Lazard v16, IRENA.
   Все расчёты — в числах с плавающей точкой, без округления.
   Единицы: рубли (₽), МВт·ч, млн ₽ для агрегатов.
   ============================================================ */

function econFormReset(){
  /* Полный сброс всех полей формы */
  var fields = ["econPower","econYield","econCapex","econOpex","econTariff",
                "econDegr","econDiscount","econLife","econTariffGrowth","econOpexGrowth",
                "econTax","econSubsidy","econDebtPct","econDebtRate","econEquityRate"];
  var defaults = {
    econPower: ECON_DEFAULTS.power,
    econYield: ECON_DEFAULTS.yieldPerKW,
    econCapex: ECON_DEFAULTS.capex,
    econOpex: ECON_DEFAULTS.opexPct,
    econTariff: ECON_DEFAULTS.tariff,
    econDegr: ECON_DEFAULTS.degr,
    econDiscount: ECON_DEFAULTS.discount,
    econLife: ECON_DEFAULTS.life,
    econTariffGrowth: ECON_DEFAULTS.tariffGrowth,
    econOpexGrowth: ECON_DEFAULTS.opexGrowth,
    econTax: ECON_DEFAULTS.tax,
    econSubsidy: ECON_DEFAULTS.subsidy,
    econDebtPct: ECON_DEFAULTS.debtPct,
    econDebtRate: ECON_DEFAULTS.debtRate,
    econEquityRate: ECON_DEFAULTS.equityRate
  };
  fields.forEach(function(id){
    var el=$(id);
    if(el && defaults[id]!==undefined) el.value = defaults[id];
  });
  if($("econDegrType")) $("econDegrType").value = ECON_DEFAULTS.degrType;
  if($("econWaccMode")) $("econWaccMode").checked = ECON_DEFAULTS.waccMode;
  if($("econRegion")){
    /* сохранить первый регион, если был */
    var firstOpt = $("econRegion").options[0];
    if(firstOpt) $("econRegion").value = firstOpt.value;
  }
}

function calcEconomics(){
  var p = {
    region:         $("econRegion") ? $("econRegion").value : "",
    power:          Number($("econPower").value)         || 0,   // МВт
    yieldPerKW:     Number($("econYield").value)         || 0,   // кВт·ч/кВт·год
    capex:          Number($("econCapex").value)         || 0,   // млн ₽
    opexPct:        Number($("econOpex").value)          || 0,   // % от CAPEX/год
    tariff:         Number($("econTariff").value)        || 0,   // ₽/кВт·ч
    degrType:       ($("econDegrType") && $("econDegrType").value) || "linear",
    degr:           (Number($("econDegr").value)  || 0) / 100,   // доля/год
    discount:       (Number($("econDiscount").value) || 0) / 100, // доля
    life:           Math.max(5, Number($("econLife").value) || 25),
    tariffGrowth:   (Number($("econTariffGrowth").value) || 0) / 100,
    opexGrowth:     (Number($("econOpexGrowth").value)   || 0) / 100,
    tax:            (Number($("econTax").value)          || 0) / 100,
    subsidy:        Number($("econSubsidy").value)       || 0,    // млн ₽/год
    waccMode:       !!($("econWaccMode") && $("econWaccMode").checked),
    debtPct:        (Number($("econDebtPct").value)      || 0) / 100,
    debtRate:       (Number($("econDebtRate").value)     || 0) / 100,
    equityRate:     (Number($("econEquityRate").value)   || 0) / 100
  };

  /* ----- Ставка дисконтирования (WACC если включён) ----- */
  if(p.waccMode){
    var wacc = p.debtPct * p.debtRate * (1 - p.tax) + (1 - p.debtPct) * p.equityRate;
    p.discount = wacc;
    if($("econDiscount")) $("econDiscount").value = +(wacc*100).toFixed(2);
  }

  /* ----- Валидация ----- */
  var errs = [];
  if(p.power <= 0)   errs.push("Мощность должна быть > 0 МВт");
  if(p.yieldPerKW <= 0) errs.push("Удельная выработка должна быть > 0 кВт·ч/кВт·год");
  if(p.capex <= 0)   errs.push("CAPEX должен быть > 0 млн ₽");
  if(p.tariff <= 0)  errs.push("Тариф должен быть > 0 ₽/кВт·ч");
  if(p.life < 5 || p.life > 40) errs.push("Срок жизни — от 5 до 40 лет");
  if(p.discount < 0 || p.discount > 1) errs.push("Ставка дисконтирования — 0…100%");
  if(p.degr < 0 || p.degr > 0.2) errs.push("Деградация — 0…20%/год");
  if(p.tariffGrowth < -0.5 || p.tariffGrowth > 0.5) errs.push("Рост тарифа — -50…50%/год");
  if(p.opexGrowth < -0.5 || p.opexGrowth > 0.5) errs.push("Рост OPEX — -50…50%/год");
  if(errs.length){
    showEconEmpty();
    showMsg("econMsg", errs.join("; "), "err");
    return;
  }
  if($("econMsg")){ $("econMsg").className="msg"; $("econMsg").textContent=""; }

  var N = Math.round(p.life);
  var CAPEX_RUB = p.capex * 1e6;            // ₽
  var opex0_RUB = CAPEX_RUB * p.opexPct/100; // ₽/год, год 1

  /* ----- По годам: генерация, выручка, OPEX, EBITDA, налог, CF, DCF ----- */
  var rows = [];
  var cumCF = 0, cumDCF = 0;
  var paybackYear = null, dPayback = null;
  var totalGen_kWh = 0, totalRev = 0, totalOpex = 0, totalTax = 0, totalSub = 0, totalCF = 0;
  var cfs = [], dcf = [];
  for(var t = 1; t <= N; t++){
    /* Выработка: экспоненциальная или линейная деградация */
    var degrFactor = (p.degrType === "exp")
      ? Math.pow(1 - p.degr, t - 1)          // экспоненциальная (компаундная)
      : Math.max(0, 1 - p.degr * (t - 1));   // линейная
    var gen_kWh = p.power * 1000 * p.yieldPerKW * degrFactor; // кВт·ч/год

    /* Тариф растёт с tariffGrowth */
    var price = p.tariff * Math.pow(1 + p.tariffGrowth, t - 1); // ₽/кВт·ч
    var rev_RUB = gen_kWh * price;

    /* OPEX растёт с opexGrowth */
    var opex_RUB = opex0_RUB * Math.pow(1 + p.opexGrowth, t - 1);

    /* Субсидия в млн ₽/год → ₽ */
    var sub_RUB = p.subsidy * 1e6;

    /* EBITDA (не учитывает амортизацию и проценты — упрощённо) */
    var ebitda = rev_RUB - opex_RUB + sub_RUB;

    /* Налог на прибыль: только если EBITDA > 0; отрицательный EBITDA даёт налоговый щит 0 (упрощённо) */
    var tax = Math.max(0, ebitda * p.tax);

    /* Чистый денежный поток (упрощённо, без учёта долга) */
    var cf = ebitda - tax;

    /* Дисконтирование */
    var df = Math.pow(1 + p.discount, -t);
    var dfc = cf * df;

    cumCF += cf;
    cumDCF += dfc;

    /* Простой и дисконтированный payback — с линейной интерполяцией внутри года */
    if(paybackYear === null && cumCF >= CAPEX_RUB){
      var over = cumCF - CAPEX_RUB;
      paybackYear = t - over / cf; // точнее, чем «t»
    }
    if(dPayback === null && cumDCF >= CAPEX_RUB){
      var overD = cumDCF - CAPEX_RUB;
      dPayback = t - overD / dfc;
    }

    totalGen_kWh += gen_kWh;
    totalRev += rev_RUB;
    totalOpex += opex_RUB;
    totalTax += tax;
    totalSub += sub_RUB;
    totalCF += cf;
    cfs.push(cf); dcf.push(dfc);

    rows.push({
      year: t,
      degrFactor: degrFactor,
      gen_kWh: gen_kWh,
      price: price,
      rev_RUB: rev_RUB,
      opex_RUB: opex_RUB,
      subsidy_RUB: sub_RUB,
      ebitda: ebitda,
      tax: tax,
      cf: cf,
      df: df,
      dfc: dfc,
      cumCF: cumCF,
      cumDCF: cumDCF
    });
  }

  var NPV = cumDCF - CAPEX_RUB;

  /* ----- IRR (внутренняя норма доходности) ----- */
  function npvAt(r){
    var v = -CAPEX_RUB;
    for(var i = 0; i < N; i++) v += cfs[i] / Math.pow(1 + r, i + 1);
    return v;
  }
  var irr = null;
  if(NPV > 0){
    /* расширенный поиск IRR: -99% … +1000% */
    var lo = -0.99, hi = 10;
    /* проверяем, что на границах разные знаки */
    if(npvAt(hi) < 0 && npvAt(lo) > 0){
      for(var it = 0; it < 200; it++){
        var mid = (lo + hi) / 2;
        if(npvAt(mid) > 0) lo = mid; else hi = mid;
      }
      irr = (lo + hi) / 2;
    }
  }

  /* ----- MIRR (модифицированная IRR) ----- */
  var mirr = null;
  if(p.discount > 0 && N > 0){
    var fvNeg = 0, pvPos = 0;
    for(var i2 = 0; i2 < N; i2++){
      if(cfs[i2] < 0){
        fvNeg += cfs[i2] * Math.pow(1 + p.discount, N - i2);
      } else {
        pvPos += cfs[i2] / Math.pow(1 + p.discount, i2 + 1);
      }
    }
    /* +CAPEX как отрицательный CF в год 0 */
    fvNeg += -CAPEX_RUB * Math.pow(1 + p.discount, N);
    if(pvPos > 0 && fvNeg < 0){
      mirr = Math.pow(pvPos / -fvNeg, 1/N) - 1;
    }
  }

  /* ----- LCOE (levelized, IEA) ----- */
  var lcoe = null;
  if(totalGen_kWh > 0){
    var pvCost = CAPEX_RUB;
    for(var i3 = 0; i3 < N; i3++){
      pvCost += rows[i3].opex_RUB / Math.pow(1 + p.discount, i3 + 1);
    }
    var pvGen = 0;
    for(var i4 = 0; i4 < N; i4++){
      pvGen += rows[i4].gen_kWh / Math.pow(1 + p.discount, i4 + 1);
    }
    lcoe = pvGen > 0 ? pvCost / pvGen : null;
  }

  /* ----- PI (profitability index) ----- */
  var PI = CAPEX_RUB > 0 ? (NPV + CAPEX_RUB) / CAPEX_RUB : null;

  /* ----- Среднегодовой CF, макс. дефицит ----- */
  var avgCF = N > 0 ? totalCF / N : null;
  var minCum = 0;
  rows.forEach(function(r){ if(r.cumCF < minCum) minCum = r.cumCF; });

  /* ----- Вердикт ----- */
  var verdict, verdictCls;
  if(NPV > 0 && PI != null && PI > 1.1 && (irr == null || irr > p.discount)){
    verdict = "Проект целесообразен"; verdictCls = "ok";
  } else if(NPV > 0){
    verdict = "Проект на грани рентабельности"; verdictCls = "warn";
  } else if(NPV > -0.1 * CAPEX_RUB){
    verdict = "Проект убыточен, но близок к окупаемости"; verdictCls = "warn";
  } else {
    verdict = "Проект нецелесообразен"; verdictCls = "err";
  }

  /* ============ ВЫВОД: ГЛАВНЫЙ РЕЗУЛЬТАТ ============ */
  var pbEl = $("econPayback");
  if(pbEl){
    pbEl.textContent = paybackYear != null
      ? fmtNum(paybackYear, 2) + " лет"
      : "не окупается за " + N + " лет";
  }
  var pbSub = $("econPaybackSub");
  if(pbSub){
    pbSub.textContent = (dPayback != null)
      ? "Дисконтированный: " + fmtNum(dPayback, 2) + " лет · NPV: " + fmtMoney(NPV/1e6, 2)
      : "Дисконтированный: не окупается · NPV: " + fmtMoney(NPV/1e6, 2);
  }
  var resEl = $("econResult");
  if(resEl) resEl.className = "econ-result " + verdictCls;

  /* ============ KPI (расширенный набор) ============ */
  var kpiHtml = "";
  function kpi(label, val, sub, cls){
    return '<div class="kpi"><div class="label">'+label+'</div>'+
           '<div class="val '+(cls||'')+'">'+val+'</div>'+
           (sub?'<div class="delta">'+sub+'</div>':'')+'</div>';
  }
  kpiHtml += kpi("NPV", fmtMoney(NPV/1e6, 2), verdict, verdictCls==="ok"?"pos":verdictCls==="err"?"neg":"");
  kpiHtml += kpi("IRR", irr != null ? fmtPct(irr*100, 2) : "—", "внутренняя норма");
  kpiHtml += kpi("MIRR", mirr != null ? fmtPct(mirr*100, 2) : "—", "модифицированная");
  kpiHtml += kpi("PI", PI != null ? fmtNum(PI, 3) : "—", PI > 1 ? "индекс > 1 — OK" : "индекс < 1");
  kpiHtml += kpi("LCOE", lcoe != null ? fmtNum(lcoe, 3) + " ₽/кВт·ч" : "—", "себестоимость");
  kpiHtml += kpi("CAPEX", fmtMoney(p.capex, 2), "инвестиции");
  kpiHtml += kpi("Год 1", fmtNum(rows[0].gen_kWh / 1000, 1) + " МВт·ч", "выработка");
  kpiHtml += kpi("Σ CF", fmtMoney(totalCF/1e6, 2), "за " + N + " лет");
  kpiHtml += kpi("Σ Выручка", fmtMoney(totalRev/1e6, 2), "за " + N + " лет");
  kpiHtml += kpi("Σ OPEX", fmtMoney(totalOpex/1e6, 2), "за " + N + " лет");
  kpiHtml += kpi("Σ Налог", fmtMoney(totalTax/1e6, 2), "налог на прибыль");
  kpiHtml += kpi("Ставка диск.", fmtPct(p.discount*100, 2), p.waccMode ? "WACC" : "задана");
  kpiHtml += kpi("Ср. CF/год", fmtMoney(avgCF/1e6, 2), "среднегодовой");
  if($("econKpis")) $("econKpis").innerHTML = kpiHtml;

  /* ============ ГРАФИКИ ============ */
  if(!plotlyReady()) return;

  var yearsX = rows.map(function(r){return r.year;});

  /* 1. Накопленный CF + CAPEX + точка окупаемости */
  var cumPlot = [0].concat(rows.map(function(r){ return r.cumCF/1e6; }));
  var cumPlotX = [0].concat(yearsX);
  var capexLine = cumPlotX.map(function(){ return CAPEX_RUB/1e6; });
  var traces1 = [
    {name:"Накопленный CF", x:cumPlotX, y:cumPlot, type:"scatter", mode:"lines+markers",
     line:{width:2.6, color:PALETTE[0]}, marker:{size:5, color:PALETTE[0]}},
    {name:"CAPEX", x:cumPlotX, y:capexLine, type:"scatter", mode:"lines",
     line:{color:"#dc2626", width:1.8, dash:"dash"}}
  ];
  if(paybackYear != null){
    traces1.push({name:"Точка окупаемости", x:[paybackYear], y:[CAPEX_RUB/1e6],
      type:"scatter", mode:"markers", marker:{size:14, color:"#059669", symbol:"star", line:{color:"#fff", width:2}},
      showlegend:true});
  }
  Plotly.newPlot($("econChart1"), traces1, baseLayout({
    title:{text:"Накопительный денежный поток, млн ₽"},
    yaxis:{title:{text:"млн ₽"}, gridcolor:gridColor()},
    xaxis:{title:{text:"Год"}, gridcolor:gridColor()}
  }), PLOT_CFG);

  /* 2. Годовая выручка, OPEX, EBITDA, CF — бар + линия */
  Plotly.newPlot($("econChart2"), [
    {name:"Выручка", x:yearsX, y:rows.map(function(r){return r.rev_RUB/1e6;}), type:"bar", marker:{color:"#10b981"}},
    {name:"OPEX",    x:yearsX, y:rows.map(function(r){return r.opex_RUB/1e6;}), type:"bar", marker:{color:"#f59e0b"}},
    {name:"Налог",   x:yearsX, y:rows.map(function(r){return r.tax/1e6;}), type:"bar", marker:{color:"#dc2626"}},
    {name:"Чистый CF", x:yearsX, y:rows.map(function(r){return r.cf/1e6;}), type:"scatter", mode:"lines+markers",
      line:{width:2.4, color:PALETTE[1]}, marker:{size:6, color:PALETTE[1]}}
  ], baseLayout({
    title:{text:"Годовая выручка, OPEX, налог и чистый CF, млн ₽"},
    barmode:"group",
    yaxis:{title:{text:"млн ₽"}, gridcolor:gridColor()}
  }), PLOT_CFG);

  /* 3. DCF по годам + NPV-кривая (cumDCF − CAPEX) */
  var npvCurve = [0].concat(rows.map(function(r){ return (r.cumDCF - CAPEX_RUB)/1e6; }));
  Plotly.newPlot($("econChart3"), [
    {name:"DCF по годам", x:yearsX, y:rows.map(function(r){return r.dfc/1e6;}), type:"bar", marker:{color:"#0891b2"}},
    {name:"NPV-кривая", x:cumPlotX, y:npvCurve, type:"scatter", mode:"lines+markers",
     line:{width:2.6, color:PALETTE[0]}, marker:{size:5, color:PALETTE[0]}, yaxis:"y2"}
  ], baseLayout({
    title:{text:"Дисконтированные поступления и NPV-кривая, млн ₽"},
    yaxis:{title:{text:"DCF, млн ₽"}, gridcolor:gridColor()},
    yaxis2:{title:{text:"NPV"}, overlaying:"y", side:"right", gridcolor:"transparent"}
  }), PLOT_CFG);

  /* ============ ДЕТАЛЬНАЯ ТАБЛИЦА ============ */
  buildEconTable(rows, p, CAPEX_RUB, NPV, irr);

  /* ============ ЧУВСТВИТЕЛЬНОСТЬ ============ */
  buildEconSensitivity(p, CAPEX_RUB);

  /* ============ СОХРАНЕНИЕ ============ */
  state.econ = p;
  state.econ._results = {
    NPV: NPV, IRR: irr, MIRR: mirr, LCOE: lcoe, PI: PI,
    paybackYear: paybackYear, dPayback: dPayback,
    totalGen_kWh: totalGen_kWh, totalRev: totalRev, totalOpex: totalOpex,
    totalCF: totalCF, totalTax: totalTax, verdict: verdict
  };
  autosave();

  function kpiSafe(t){ return escapeHtml(t); }
}

function showEconEmpty(){
  ["econPayback","econPaybackSub"].forEach(function(id){
    var el = $(id);
    if(el) el.textContent = "—";
  });
  if($("econPaybackSub")) $("econPaybackSub").textContent = "Введите параметры и нажмите «Рассчитать»";
  if($("econKpis")) $("econKpis").innerHTML = "";
  if($("econTableWrap")) $("econTableWrap").innerHTML = "";
  if($("econSensWrap")) $("econSensWrap").innerHTML = "";
  if(plotlyReady()){
    ["econChart1","econChart2","econChart3"].forEach(function(id){
      if($(id)) Plotly.newPlot($(id), [], baseLayout({title:{text:"Введите параметры и нажмите «Рассчитать»"}}), PLOT_CFG);
    });
  }
}

function buildEconTable(rows, p, CAPEX_RUB, NPV, irr){
  if(!$("econTableWrap")) return;
  var html = '<div class="tbl-wrap small"><table><thead><tr>'+
    '<th>Год</th>'+
    '<th class="num">Деградация</th>'+
    '<th class="num">Выработка, МВт·ч</th>'+
    '<th class="num">Тариф, ₽/кВт·ч</th>'+
    '<th class="num">Выручка, млн ₽</th>'+
    '<th class="num">OPEX, млн ₽</th>'+
    '<th class="num">Субсидия, млн ₽</th>'+
    '<th class="num">EBITDA, млн ₽</th>'+
    '<th class="num">Налог, млн ₽</th>'+
    '<th class="num">CF, млн ₽</th>'+
    '<th class="num">DF</th>'+
    '<th class="num">DCF, млн ₽</th>'+
    '<th class="num">Накопл. CF, млн ₽</th>'+
    '<th class="num">Накопл. DCF, млн ₽</th>'+
  '</tr></thead><tbody>';
  rows.forEach(function(r){
    html += '<tr><td>'+r.year+'</td>'+
      '<td class="num">'+fmtNum(r.degrFactor*100, 2)+'%</td>'+
      '<td class="num">'+fmtNum(r.gen_kWh/1000, 1)+'</td>'+
      '<td class="num">'+fmtNum(r.price, 4)+'</td>'+
      '<td class="num">'+fmtNum(r.rev_RUB/1e6, 2)+'</td>'+
      '<td class="num">'+fmtNum(r.opex_RUB/1e6, 2)+'</td>'+
      '<td class="num">'+fmtNum(r.subsidy_RUB/1e6, 2)+'</td>'+
      '<td class="num '+(r.ebitda>0?'pos':'neg')+'">'+fmtNum(r.ebitda/1e6, 2)+'</td>'+
      '<td class="num">'+fmtNum(r.tax/1e6, 2)+'</td>'+
      '<td class="num '+(r.cf>0?'pos':'neg')+'">'+fmtNum(r.cf/1e6, 2)+'</td>'+
      '<td class="num">'+fmtNum(r.df, 4)+'</td>'+
      '<td class="num">'+fmtNum(r.dfc/1e6, 2)+'</td>'+
      '<td class="num '+(r.cumCF>0?'pos':'neg')+'">'+fmtNum(r.cumCF/1e6, 2)+'</td>'+
      '<td class="num '+(r.cumDCF>0?'pos':'neg')+'">'+fmtNum((r.cumDCF-CAPEX_RUB)/1e6, 2)+'</td></tr>';
  });
  html += '</tbody></table></div>';
  html += '<div class="row" style="margin-top:8px"><button id="btnEconExportCsv" class="ghost" type="button">📥 Экспорт таблицы CSV</button>'+
          '<button id="btnEconExportJson" class="ghost" type="button">📥 Экспорт JSON</button></div>';
  $("econTableWrap").innerHTML = html;
  on($("btnEconExportCsv"), "click", function(){ exportEconTableCsv(rows, p); });
  on($("btnEconExportJson"), "click", function(){ exportEconJson(rows, p, CAPEX_RUB, NPV, irr); });
}

function buildEconSensitivity(p, CAPEX_RUB){
  if(!$("econSensWrap")) return;
  /* Чувствительность NPV к тарифу (-30%…+30%) и ставке дисконтирования (5%…25%) */
  var tariffSteps = [-0.3,-0.2,-0.1,0,0.1,0.2,0.3];
  var discSteps = [0.05,0.08,0.10,0.12,0.15,0.18,0.20,0.25];
  var z = [], rows = [];
  discSteps.forEach(function(r){
    var row = [];
    tariffSteps.forEach(function(dt){
      var t = p.tariff * (1+dt);
      var n = Math.round(p.life);
      var npv = -CAPEX_RUB;
      var opex0 = CAPEX_RUB * p.opexPct/100;
      for(var i=1;i<=n;i++){
        var degrFactor = (p.degrType==="exp") ? Math.pow(1-p.degr, i-1) : Math.max(0, 1 - p.degr*(i-1));
        var gen = p.power*1000*p.yieldPerKW*degrFactor;
        var price = t * Math.pow(1+p.tariffGrowth, i-1);
        var rev = gen*price;
        var opex = opex0*Math.pow(1+p.opexGrowth, i-1);
        var sub = p.subsidy*1e6;
        var ebitda = rev - opex + sub;
        var tax = Math.max(0, ebitda*p.tax);
        var cf = ebitda - tax;
        npv += cf / Math.pow(1+r, i);
      }
      row.push(npv/1e6);
    });
    z.push(row);
    rows.push(r);
  });
  if(plotlyReady()){
    Plotly.newPlot($("econSensWrap"), [{
      type:"heatmap",
      z: z,
      x: tariffSteps.map(function(d){ return (d*100).toFixed(0)+"%"; }),
      y: discSteps.map(function(r){ return (r*100).toFixed(0)+"%"; }),
      colorscale: [[0,"#dc2626"],[0.5,"#f59e0b"],[1,"#059669"]],
      colorbar:{title:"NPV, млн ₽"},
      hovertemplate: "Ставка: %{y}<br>Тариф: %{x}<br>NPV: %{z:.1f} млн ₽<extra></extra>"
    }], baseLayout({
      title:{text:"Чувствительность NPV к тарифу и ставке дисконтирования"},
      xaxis:{title:{text:"Изменение тарифа"}, gridcolor:gridColor()},
      yaxis:{title:{text:"Ставка дисконтирования"}, gridcolor:gridColor()}
    }), PLOT_CFG);
  }
}

function exportEconTableCsv(rows, p){
  var out = [["Год","Деградация,%","Выработка,МВт·ч","Тариф,₽/кВт·ч","Выручка,млн ₽","OPEX,млн ₽","Субсидия,млн ₽","EBITDA,млн ₽","Налог,млн ₽","CF,млн ₽","DF","DCF,млн ₽","Накопл.CF,млн ₽","Накопл.DCF,млн ₽"]];
  rows.forEach(function(r){
    out.push([r.year, (r.degrFactor*100).toFixed(2), (r.gen_kWh/1000).toFixed(2), r.price.toFixed(4),
      (r.rev_RUB/1e6).toFixed(2), (r.opex_RUB/1e6).toFixed(2), (r.subsidy_RUB/1e6).toFixed(2),
      (r.ebitda/1e6).toFixed(2), (r.tax/1e6).toFixed(2), (r.cf/1e6).toFixed(2),
      r.df.toFixed(4), (r.dfc/1e6).toFixed(2), (r.cumCF/1e6).toFixed(2), ((r.cumDCF)/1e6).toFixed(2)]);
  });
  var csv = "\uFEFF" + out.map(function(r){return r.join(";");}).join("\n");
  downloadBlob(new Blob([csv], {type:"text/csv;charset=utf-8"}), "экономика_СЭС_"+p.region.replace(/[\\/:*?"<>|]/g,"_")+".csv");
}

function exportEconJson(rows, p, CAPEX_RUB, NPV, irr){
  var payload = {
    parameters: p,
    capex_rub: CAPEX_RUB,
    npv_rub: NPV,
    irr: irr,
    rows: rows
  };
  downloadBlob(new Blob([JSON.stringify(payload, null, 2)], {type:"application/json;charset=utf-8"}),
    "экономика_СЭС_"+p.region.replace(/[\\/:*?"<>|]/g,"_")+".json");
}

/* ============================================================
   ЭКСПОРТ
   ============================================================ */
function fileBase(){ var d=new Date(); var p=function(n){return String(n).padStart(2,"0");}; return "СЭС_РФ_"+d.getFullYear()+"-"+p(d.getMonth()+1)+"-"+p(d.getDate()); }
function downloadBlob(blob,filename){
  try{
    var url=URL.createObjectURL(blob);
    var a=document.createElement("a"); a.href=url; a.download=filename;
    document.body.appendChild(a); a.click();
    setTimeout(function(){URL.revokeObjectURL(url);a.remove();},200);
    toast("Скачано: "+filename,"ok"); tInfo("Скачан: "+filename);
  }catch(e){ tErr("Скачивание: "+e.message); }
}
function buildExportTraces(){
  var sel=REGIONS.filter(function(r){return state.selected[r];});
  var years=yearsAll(), type=$("chartType").value;
  if(type==="stack100"){
    var totals={};
    years.forEach(function(y){ var t=0; sel.forEach(function(r){ var v=valueAt(r,y); if(v!=null) t+=v; }); totals[y]=t; });
    return sel.map(function(r,i){
      var ys=years.map(function(y){ var v=valueAt(r,y), t=totals[y]; return (v!=null&&t>0)?v/t*100:null; });
      return {name:r, x:years, y:ys, type:"bar", marker:{color:PALETTE[i%PALETTE.length]}};
    });
  }
  return sel.map(function(r,i){
    var ys=years.map(function(y){ var v=valueAt(r,y); return v!=null?convertUnits(v):null; });
    var color=PALETTE[i%PALETTE.length];
    if(type==="lines") return {name:r, x:years, y:ys, type:"scatter", mode:"lines+markers", line:{width:2.4,color:color}, marker:{size:7,color:color}};
    if(type==="area") return {name:r, x:years, y:ys, type:"scatter", mode:"lines+markers", fill:"tozeroy", line:{width:2,color:color}, marker:{size:6,color:color}};
    return {name:r, x:years, y:ys, type:"bar", marker:{color:color}};
  });
}
function buildExportChart(divId,width,height){
  return new Promise(function(resolve,reject){
    if(!plotlyReady()){ reject(new Error("Plotly не загружен")); return; }
    var stage=$("exportStage"); stage.innerHTML="";
    var tmp=document.createElement("div"); tmp.id=divId;
    tmp.style.width=(width||1600)+"px"; tmp.style.height=(height||900)+"px"; tmp.style.background="#ffffff";
    stage.appendChild(tmp);
    var layout={
      title:{text:"Выработка СЭС по регионам РФ, "+unitLabel(), font:{size:22,family:'"Times New Roman"'}},
      xaxis:{title:{text:"Год",font:{size:16}}, gridcolor:"#e1e7f2"},
      yaxis:{title:{text:unitLabel(),font:{size:16}}, gridcolor:"#e1e7f2"},
      legend:{orientation:"h", y:-0.18, font:{size:14}},
      paper_bgcolor:"#ffffff", plot_bgcolor:"#ffffff",
      font:{family:'"Times New Roman", Georgia, serif', color:"#0f1a2b", size:14},
      margin:{l:90,r:40,t:70,b:90}, hovermode:"x unified"
    };
    Plotly.newPlot(tmp,buildExportTraces(),layout,{displayModeBar:false, staticPlot:true});
    Plotly.toImage(tmp,{format:"png", width:width||1600, height:height||900}).then(function(url){
      stage.innerHTML=""; resolve(url);
    }).catch(function(e){ stage.innerHTML=""; reject(e); });
  });
}
function exportRaster(fmt){
  return buildExportChart("exportChartTmp",1600,900).then(function(u){
    return fetch(u).then(function(r){return r.blob();});
  }).then(function(b){ downloadBlob(b, fileBase()+"."+fmt); });
}
function exportSvg(){
  return buildExportChart("exportChartTmp",1600,900).then(function(){
    var stage=$("exportStage"); stage.innerHTML="";
    var tmp=document.createElement("div"); tmp.id="exportChartTmp";
    tmp.style.width="1600px"; tmp.style.height="900px"; tmp.style.background="#fff";
    stage.appendChild(tmp);
    var layout={
      title:{text:"Выработка СЭС по регионам РФ, "+unitLabel(), font:{size:22}},
      xaxis:{title:{text:"Год"}}, yaxis:{title:{text:unitLabel()}},
      legend:{orientation:"h", y:-0.18}, paper_bgcolor:"#fff", plot_bgcolor:"#fff",
      font:{family:'"Times New Roman", Georgia, serif', color:"#0f1a2b", size:14},
      margin:{l:90,r:40,t:70,b:90}
    };
    Plotly.newPlot(tmp,buildExportTraces(),layout,{displayModeBar:false, staticPlot:true});
    return Plotly.toImage(tmp,{format:"svg", width:1600, height:900}).then(function(u){
      stage.innerHTML="";
      return fetch(u).then(function(r){return r.blob();}).then(function(b){ downloadBlob(b, fileBase()+".svg"); });
    });
  });
}
function buildReportHtml(imgDataUrl,mode){
  mode=mode||"short";
  var stats=buildStats(), years=yearsAll(), tot=totalsByYearSelected();
  var cap=($("figCaption")&&$("figCaption").value)||"Выработка СЭС";
  var src=($("figSource")&&$("figSource").value)||"";
  var org=($("sigOrg")&&$("sigOrg").value)||"";
  var auth=($("sigAuthor")&&$("sigAuthor").value)||"";
  var date=($("sigDate")&&$("sigDate").value)||"";
  var note=($("sigNote")&&$("sigNote").value)||"";
  var city=($("sigCity")&&$("sigCity").value)||"Москва";
  var statsTable=stats.length?(
    '<h2>2. Сводная таблица</h2><table class="tbl"><thead><tr><th>Регион</th><th>N</th><th>Σ</th><th>Ср.</th><th>Мед.</th><th>σ</th><th>CV%</th><th>Min</th><th>Max</th><th>CAGR%</th><th>Наклон</th><th>R²</th><th>Доля%</th></tr></thead><tbody>'+
    stats.map(function(r){
      return '<tr><td>'+escapeHtml(r.region)+'</td><td class="num">'+r.n+'</td>'+
        '<td class="num">'+fmtNum(convertUnits(r.sum),1)+'</td>'+
        '<td class="num">'+fmtNum(convertUnits(r.mean),2)+'</td>'+
        '<td class="num">'+fmtNum(convertUnits(r.median),2)+'</td>'+
        '<td class="num">'+fmtNum(r.std,2)+'</td>'+
        '<td class="num">'+(r.cv!=null?fmtNum(r.cv,2):"—")+'</td>'+
        '<td class="num">'+fmtNum(convertUnits(r.min),2)+'</td>'+
        '<td class="num">'+fmtNum(convertUnits(r.max),2)+'</td>'+
        '<td class="num">'+(r.cagr!=null?fmtNum(r.cagr,2):"—")+'</td>'+
        '<td class="num">'+(r.slope!=null?fmtNum(r.slope,3):"—")+'</td>'+
        '<td class="num">'+(r.r2!=null?fmtNum(r.r2,3):"—")+'</td>'+
        '<td class="num">'+(r.share!=null?fmtNum(r.share,2):"—")+'</td></tr>';
    }).join('')+'</tbody></table>'
  ):'<p>Нет данных.</p>';
  var dynTable=years.length?(
    '<h2>3. Динамика</h2><table class="tbl"><thead><tr><th>Год</th><th>Сумма</th></tr></thead><tbody>'+
    years.map(function(y){ return '<tr><td>'+y+'</td><td class="num">'+(tot[y]!=null?fmtNum(convertUnits(tot[y]),2):"—")+'</td></tr>'; }).join('')+
    '</tbody></table>'
  ):'';
  var methodBlock="";
  if(mode==="full"){
    methodBlock='<h2>Методика</h2><p><b>Единица:</b> '+unitLabel()+'.</p><p>Среднее, медиана, σ, CV, CAGR, тренды, HHI, Джини, z-score, IQR, MAD, Манн-Кендалл, PCA, k-means, NPV, IRR, MIRR, LCOE, PI — стандартные методы.</p><p>Полные формулы — <a href="methodology.html">methodology.html</a>.</p>';
  }
  var dataRows=state.data.slice().sort(function(a,b){ var c=a.region.localeCompare(b.region,"ru"); return c!==0?c:a.year-b.year; })
    .map(function(d){ return '<tr><td>'+escapeHtml(d.region)+'</td><td class="num">'+d.year+'</td><td class="num">'+fmtNum(convertUnits(d.value),2)+'</td><td>'+escapeHtml(d.status||"факт")+'</td><td>'+escapeHtml(d.note||"")+'</td></tr>'; }).join("");
  var appTable='<h2>Приложение. Исходные данные</h2><table class="tbl"><thead><tr><th>Регион</th><th>Год</th><th>Значение</th><th>Статус</th><th>Примечание</th></tr></thead><tbody>'+dataRows+'</tbody></table>';
  var css='@page{size:A4; margin:20mm 15mm 20mm 25mm}body{font-family:"Times New Roman",serif;color:#111;background:#fff;font-size:13pt;line-height:1.5;margin:0}.title{text-align:center;margin-top:80mm}.title .org{font-size:14pt;text-transform:uppercase;letter-spacing:.5px;margin-bottom:60mm}.title h1{font-size:20pt;margin:0 0 8mm 0}.title .sub{font-size:14pt;color:#333;margin-bottom:60mm}.title .author{margin-top:40mm;font-size:13pt}.title .city{margin-top:30mm;font-size:13pt}.page-break{page-break-after:always}h2{font-size:15pt;margin:18px 0 10px 0;border-bottom:1px solid #999;padding-bottom:4px}.fig{margin:20px 0;text-align:center}.fig img{max-width:100%;border:1px solid #ccc}.fig .cap{font-size:12pt;margin-top:8px;font-style:italic}.tbl{border-collapse:collapse;width:100%;font-family:"Arial",sans-serif;font-size:10pt;margin:10px 0}.tbl th,.tbl td{border:1px solid #888;padding:4px 6px;text-align:left}.tbl th{background:#eef2f7}.tbl .num{text-align:right}.sig{display:flex;justify-content:space-between;margin-top:30px;font-size:12pt}@media print{.noprint{display:none}}';
  return '<!DOCTYPE html><html lang="ru"><head><meta charset="utf-8"><title>'+escapeHtml(cap)+'</title><style>'+css+'</style></head><body>'+
    '<div class="noprint" style="position:sticky;top:0;background:#1e40af;color:#fff;padding:10px 14px;text-align:center;font-family:Arial,sans-serif;font-size:13px;z-index:9">Режим: <b>'+(mode==="full"?"полный":"краткий")+'</b>. Нажмите <b>Ctrl+P</b> → «Сохранить как PDF». <button onclick="window.print()" style="margin-left:12px;padding:6px 14px;background:#fff;color:#1e40af;border:none;border-radius:6px;cursor:pointer;font-weight:600">Печать</button></div>'+
    '<div class="title">'+(org?'<div class="org">'+escapeHtml(org)+'</div>':'')+'<h1>'+escapeHtml(cap)+'</h1><div class="sub">Аналитический отчёт по солнечной энергетике РФ</div><div class="author">Составил: '+(auth||"____")+'</div><div class="author">Дата: '+escapeHtml(date)+'</div><div class="city">'+escapeHtml(city)+'</div></div>'+
    '<div class="page-break"></div><h2>1. График</h2><div class="fig">'+(imgDataUrl?'<img src="'+imgDataUrl+'">':'')+'<div class="cap">'+escapeHtml(cap)+'</div></div>'+
    statsTable+dynTable+methodBlock+appTable+(note?'<p style="font-style:italic">'+escapeHtml(note)+'</p>':'')+
    '<div class="sig"><div>Составил: __________</div><div>Дата: '+escapeHtml(date)+'</div></div></body></html>';
}
function exportHtml(){
  return buildExportChart("exportChartTmp",1600,900).then(function(png){
    downloadBlob(new Blob([buildReportHtml(png,state.expMode)],{type:"text/html;charset=utf-8"}), fileBase()+".html");
  }).catch(function(){ downloadBlob(new Blob([buildReportHtml(null,state.expMode)],{type:"text/html;charset=utf-8"}), fileBase()+".html"); });
}
function openPrintWindow(png){
  var html=buildReportHtml(png,state.expMode);
  var w=window.open("","_blank");
  if(!w){ tErr("Всплывающее окно заблокировано"); return; }
  w.document.open(); w.document.write(html); w.document.close();
  w.onload=function(){ setTimeout(function(){ try{w.focus(); w.print();}catch(e){} }, 400); };
}
function exportPdf(){ return buildExportChart("exportChartTmp",1600,900).then(openPrintWindow).catch(function(){ openPrintWindow(null); }); }
function exportDocx(){
  if(!docxReady()) throw new Error("docx не загружен");
  return buildExportChart("exportChartTmp",1600,900).catch(function(){return null;}).then(function(dataUrl){
    var D=window.docx, cap=($("figCaption")&&$("figCaption").value)||"Выработка СЭС";
    var children=[];
    var org=($("sigOrg")&&$("sigOrg").value)||"";
    if(org) children.push(new D.Paragraph({text:org.toUpperCase(), alignment:D.AlignmentType.CENTER, spacing:{after:2400}}));
    children.push(new D.Paragraph({text:cap, heading:D.HeadingLevel.HEADING_1, alignment:D.AlignmentType.CENTER}));
    children.push(new D.Paragraph({text:"Аналитический отчёт по солнечной энергетике РФ", alignment:D.AlignmentType.CENTER, spacing:{after:2400}}));
    children.push(new D.Paragraph({children:[new D.PageBreak()]}));
    children.push(new D.Paragraph({text:"1. График", heading:D.HeadingLevel.HEADING_2}));
    if(dataUrl){
      var b64=dataUrl.split(",")[1], bin=atob(b64), arr=new Uint8Array(bin.length);
      for(var i=0;i<bin.length;i++) arr[i]=bin.charCodeAt(i);
      children.push(new D.Paragraph({alignment:D.AlignmentType.CENTER, children:[new D.ImageRun({data:arr, transformation:{width:640,height:360}})]}));
    }
    var stats=buildStats();
    if(stats.length){
      children.push(new D.Paragraph({text:"2. Сводная таблица", heading:D.HeadingLevel.HEADING_2, spacing:{before:240}}));
      var hdr=["Регион","N","Σ","Ср.","Мед.","σ","CV%","Min","Max","CAGR%","Доля%"];
      var hr=new D.TableRow({tableHeader:true, children:hdr.map(function(t){ return new D.TableCell({shading:{fill:"EEF2F7"}, children:[new D.Paragraph({children:[new D.TextRun({text:t,bold:true,size:16})]})]}); })});
      var dr=stats.map(function(r){
        var vals=[r.region, r.n, fmtNum(convertUnits(r.sum),1), fmtNum(convertUnits(r.mean),2), fmtNum(convertUnits(r.median),2), fmtNum(r.std,2), r.cv!=null?fmtNum(r.cv,2):"—", fmtNum(convertUnits(r.min),2), fmtNum(convertUnits(r.max),2), r.cagr!=null?fmtNum(r.cagr,2):"—", r.share!=null?fmtNum(r.share,2):"—"];
        return new D.TableRow({children: vals.map(function(v){ return new D.TableCell({children:[new D.Paragraph({children:[new D.TextRun({text:String(v),size:16})]})]}); })});
      });
      children.push(new D.Table({rows:[hr].concat(dr), width:{size:100,type:D.WidthType.PERCENTAGE}}));
    }
    var doc=new D.Document({sections:[{children:children}]});
    return D.Packer.toBlob(doc).then(function(b){ downloadBlob(b, fileBase()+".docx"); });
  });
}
function exportCsv(){
  var rows=[["Регион","Год","Значение","Статус","Примечание"]];
  state.data.slice().sort(function(a,b){ var c=a.region.localeCompare(b.region,"ru"); return c!==0?c:a.year-b.year; })
    .forEach(function(d){ rows.push([d.region, d.year, String(d.value).replace(".",","), d.status||"факт", d.note||""]); });
  var csv="\uFEFF"+rows.map(function(r){return r.join(";");}).join("\n");
  downloadBlob(new Blob([csv],{type:"text/csv;charset=utf-8"}), fileBase()+".csv");
}
function exportJson(){
  var payload={data:state.data, selected:state.selected, stats:buildStats(), years:yearsAll(), econ:state.econ};
  downloadBlob(new Blob([JSON.stringify(payload,null,2)],{type:"application/json;charset=utf-8"}), fileBase()+".json");
}
function exportXlsx(){
  if(!xlsxReady()) throw new Error("SheetJS не загружен");
  var wb=XLSX.utils.book_new();
  var r1=[["Регион","Год","Значение","Статус","Примечание"]];
  state.data.forEach(function(d){ r1.push([d.region,d.year,d.value,d.status||"факт",d.note||""]); });
  var ws1=XLSX.utils.aoa_to_sheet(r1); ws1["!cols"]=[{wch:32},{wch:8},{wch:18},{wch:12},{wch:30}];
  XLSX.utils.book_append_sheet(wb, ws1, "Данные");
  var st=buildStats();
  var r2=[["Регион","N","Σ","Ср.","Мед.","σ","CV%","Min","Max","CAGR%","Наклон","R²","Доля%"]];
  st.forEach(function(r){ r2.push([r.region, r.n, r.sum, r.mean, r.median, r.std, r.cv, r.min, r.max, r.cagr, r.slope, r.r2, r.share]); });
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(r2), "Сводная");
  XLSX.writeFile(wb, fileBase()+".xlsx"); toast("XLSX скачан","ok");
}
function exportMd(){
  var stats=buildStats();
  var md="# Выработка СЭС по регионам РФ\n\n## Сводная\n\n| Регион | N | Σ | Ср. | CV% | CAGR% |\n|---|---:|---:|---:|---:|---:|\n";
  stats.forEach(function(r){
    md+="| "+r.region+" | "+r.n+" | "+fmtNum(convertUnits(r.sum),1)+" | "+fmtNum(convertUnits(r.mean),2)+" | "+(r.cv!=null?fmtNum(r.cv,2):"—")+" | "+(r.cagr!=null?fmtNum(r.cagr,2):"—")+" |\n";
  });
  downloadBlob(new Blob([md],{type:"text/markdown;charset=utf-8"}), fileBase()+".md");
}
function exportZip(){
  if(!jszipReady()) throw new Error("JSZip не загружен");
  var zip=new JSZip();
  var sel=REGIONS.filter(function(r){return state.selected[r];});
  sel.forEach(function(reg){
    var s=seriesFor(reg); if(!s.values.length) return;
    var lines=[["Год","Значение","Статус","Примечание"]];
    s.rows.forEach(function(row){ lines.push([row.year,row.value,row.status||"факт",row.note||""]); });
    var csv="\uFEFF"+lines.map(function(r){return r.join(";");}).join("\n");
    zip.file(reg.replace(/[\\/:*?"<>|]/g,"_")+".csv", csv);
  });
  return zip.generateAsync({type:"blob"}).then(function(b){ downloadBlob(b, fileBase()+"_по_регионам.zip"); });
}
function exportData(fmt){
  if(!state.data.length) throw new Error("Нет данных.");
  var cnt=0; for(var r in state.selected) if(state.selected[r]) cnt++;
  if(!cnt) throw new Error("Выберите хотя бы один регион.");
  if(fmt==="png"||fmt==="jpeg"||fmt==="webp") return exportRaster(fmt);
  if(fmt==="svg") return exportSvg();
  if(fmt==="pdf") return exportPdf();
  if(fmt==="html") return exportHtml();
  if(fmt==="csv") return exportCsv();
  if(fmt==="xlsx") return exportXlsx();
  if(fmt==="json") return exportJson();
  if(fmt==="docx") return exportDocx();
  if(fmt==="md") return exportMd();
  if(fmt==="zip") return exportZip();
  throw new Error("Неизвестный формат: "+fmt);
}

/* ---------- PAGES / HISTORY ---------- */
function goTo(page){
  if(!page) page="home";
  state.currentPage=page;
  qsa(".page").forEach(function(p){ p.classList.toggle("active", p.id==="page-"+page); });
  qsa(".nav a, .mobile-nav a, .drawer a").forEach(function(a){ a.classList.toggle("active", a.getAttribute("data-page")===page); });
  closeDrawer();
  var crumbMap={
    home:"Главная",data:"1.1 Ввод данных СЭС",health:"1.2 Качество данных",deepseek:"1.3 Извлечение из документов",
    view:"2.1 Графики выработки",compare:"2.2 Сравнение регионов",stats:"2.3 Статистика",trend:"2.4 Тренды и прогнозы",
    advanced:"2.5 Продвинутая аналитика",structure:"2.6 Структура рынка",anomaly:"2.7 Аномалии",heatmap:"2.8 Тепловые карты",
    region:"2.9 Паспорт региона",scenario:"2.10 Сценарии",
    economics:"3.1 Окупаемость СЭС",
    official:"4.1 Оформление отчёта",export:"4.2 Экспорт файлов",settings:"4.3 Настройки"
  };
  if($("crumb")) $("crumb").textContent = crumbMap[page]||"";
  try{ history.replaceState({page:page},"", "#"+page); }catch(e){}
  setTimeout(function(){
    try{
      if(page==="view"&&plotlyReady()){ renderMainChart(); Plotly.Plots.resize($("chart")); }
      if(page==="compare"&&plotlyReady()){ renderCompare(); Plotly.Plots.resize($("compareChart")); }
      if(page==="trend"&&plotlyReady()){ renderTrend(); Plotly.Plots.resize($("trendChart")); }
      if(page==="advanced"&&plotlyReady()) renderAdvanced();
      if(page==="structure"&&plotlyReady()){ renderStructure(); Plotly.Plots.resize($("structureChart")); }
      if(page==="anomaly") renderAnomaly();
      if(page==="heatmap"&&plotlyReady()) renderHeatmap();
      if(page==="stats") renderStats();
      if(page==="health") renderHealth();
      if(page==="region") renderRegionPassport();
      if(page==="scenario") renderScenario();
      if(page==="economics"){
        if(state.econ && state.econ.power){
          /* восстановить поля из state.econ и пересчитать */
          if($("econPower")) $("econPower").value = state.econ.power;
          if($("econYield")) $("econYield").value = state.econ.yieldPerKW;
          if($("econCapex")) $("econCapex").value = state.econ.capex;
          if($("econOpex")) $("econOpex").value = state.econ.opexPct;
          if($("econTariff")) $("econTariff").value = state.econ.tariff;
          if($("econDegr")) $("econDegr").value = state.econ.degr*100;
          if($("econDiscount")) $("econDiscount").value = state.econ.discount*100;
          if($("econLife")) $("econLife").value = state.econ.life;
          if($("econTariffGrowth")) $("econTariffGrowth").value = state.econ.tariffGrowth*100;
          if($("econOpexGrowth")) $("econOpexGrowth").value = state.econ.opexGrowth*100;
          if($("econTax")) $("econTax").value = state.econ.tax*100;
          if($("econSubsidy")) $("econSubsidy").value = state.econ.subsidy;
          if($("econDegrType")) $("econDegrType").value = state.econ.degrType;
          if($("econRegion") && state.econ.region) $("econRegion").value = state.econ.region;
          if($("econWaccMode")) $("econWaccMode").checked = state.econ.waccMode;
          if($("econDebtPct")) $("econDebtPct").value = state.econ.debtPct*100;
          if($("econDebtRate")) $("econDebtRate").value = state.econ.debtRate*100;
          if($("econEquityRate")) $("econEquityRate").value = state.econ.equityRate*100;
          calcEconomics();
        } else {
          showEconEmpty();
        }
      }
    }catch(e){ tErr("Nav: "+e.message); }
  },30);
  window.scrollTo(0,0);
  autosave();
}

function refreshAll(){
  try{ renderDataTable(); }catch(e){ tErr("Table: "+e.message); }
  try{ renderAppendTable(); }catch(e){}
  try{ renderStats(); }catch(e){}
  try{ populateRegionSelects(); }catch(e){}
  try{ if($("page-view").classList.contains("active")) renderMainChart(); }catch(e){}
  try{ if($("page-compare").classList.contains("active")) renderCompare(); }catch(e){}
  try{ if($("page-trend").classList.contains("active")) renderTrend(); }catch(e){}
  try{ if($("page-advanced").classList.contains("active")) renderAdvanced(); }catch(e){}
  try{ if($("page-structure").classList.contains("active")) renderStructure(); }catch(e){}
  try{ if($("page-anomaly").classList.contains("active")) renderAnomaly(); }catch(e){}
  try{ if($("page-heatmap").classList.contains("active")) renderHeatmap(); }catch(e){}
  try{ if($("page-health").classList.contains("active")) renderHealth(); }catch(e){}
  try{ if($("page-region").classList.contains("active")) renderRegionPassport(); }catch(e){}
  try{ if($("page-scenario").classList.contains("active")) renderScenario(); }catch(e){}
  autosave();
}

/* ---------- TERMINAL ---------- */
function renderTerm(){
  if(!$("panelLog")) return;
  var fmt=function(e){
    var hm=String(e.t.getHours()).padStart(2,"0")+":"+String(e.t.getMinutes()).padStart(2,"0")+":"+String(e.t.getSeconds()).padStart(2,"0");
    var cls=e.level==="ERR"?"t-err":e.level==="WARN"?"t-warn":"t-info";
    return '<div class="line"><span class="t-muted">['+hm+']</span> <span class="'+cls+'">'+e.level+'</span> '+escapeHtml(e.msg)+'</div>';
  };
  $("panelLog").innerHTML=LOG.info.slice(-500).map(fmt).join("")||'<div class="t-muted">— пусто —</div>';
  $("panelErr").innerHTML=LOG.err.slice(-200).map(fmt).join("")||'<div class="t-muted">— пусто —</div>';
  $("panelWarn").innerHTML=LOG.warn.slice(-200).map(fmt).join("")||'<div class="t-muted">— пусто —</div>';
  if($("panelData")){
    var years=yearsAll();
    $("panelData").innerHTML=
      '<div class="t-muted">Строк: '+state.data.length+'</div>'+
      '<div class="t-muted">Выбрано регионов: '+Object.keys(state.selected).filter(function(k){return state.selected[k];}).length+'</div>'+
      '<div class="t-muted">Годы: '+(years.length?years[0]+"–"+years[years.length-1]:"—")+'</div>'+
      '<div class="t-muted">Единица: '+unitLabel()+'</div>'+
      '<div class="t-muted">Страница: '+state.currentPage+'</div>';
  }
}

/* ---------- COMMANDS ---------- */
var COMMANDS={
  help:{cat:"Основные",desc:"Список всех команд",usage:"help",run:function(){ return "Всего: "+Object.keys(COMMANDS).length+". Вкладка «Команды»."; }},
  clear:{cat:"Данные",desc:"Очистить данные",usage:"clear",run:function(){ pushHistory(currentSnapshot()); state.data=[]; state.selected={}; refreshAll(); renderAllChips(); return "Очищено"; }},
  "load demo":{cat:"Данные",desc:"Демо-данные",usage:"load demo",run:function(){ pushHistory(currentSnapshot()); loadDemo(); return "Демо"; }},
  count:{cat:"Данные",desc:"Счётчики",usage:"count",run:function(){ return state.data.length+" строк · "+regionsWithData().length+" регионов · "+yearsAll().length+" лет"; }},
  regions:{cat:"Данные",desc:"Регионы с данными",usage:"regions",run:function(){ return regionsWithData().join("\n")||"Нет"; }},
  years:{cat:"Данные",desc:"Годы",usage:"years",run:function(){ return yearsAll().join(", ")||"Нет"; }},
  head:{cat:"Данные",desc:"Первые N",usage:"head [N]",run:function(a){ var n=Number(a[0])||10; return state.data.slice(0,n).map(function(d){return d.year+" "+d.region+": "+d.value;}).join("\n")||"Пусто"; }},
  tail:{cat:"Данные",desc:"Последние N",usage:"tail [N]",run:function(a){ var n=Number(a[0])||10; return state.data.slice(-n).map(function(d){return d.year+" "+d.region+": "+d.value;}).join("\n")||"Пусто"; }},
  missing:{cat:"Данные",desc:"Пропуски",usage:"missing",run:function(){
    var years=yearsAll(), regions=regionsWithData(), out=[];
    regions.forEach(function(r){ var s=seriesFor(r), set={}; s.years.forEach(function(y){set[y]=1;}); var gaps=years.filter(function(y){return !set[y];}); if(gaps.length) out.push(r+": "+gaps.join(", ")); });
    return out.length?out.join("\n"):"Пропусков нет";
  }},
  outliers:{cat:"Анализ",desc:"Аномалии z>2",usage:"outliers",run:function(){
    var out=[];
    regionsWithData().forEach(function(reg){
      var s=seriesFor(reg); if(s.values.length<2) return;
      var m=mean(s.values), sd=std(s.values); if(!sd) return;
      for(var i=0;i<s.values.length;i++){ var z=(s.values[i]-m)/sd; if(Math.abs(z)>2) out.push(reg+" "+s.years[i]+": z="+fmtNum(z,2)); }
    });
    return out.length?out.join("\n"):"Аномалий нет";
  }},
  top:{cat:"Анализ",desc:"Топ-5",usage:"top",run:function(){
    var yl=yearsAll(); if(!yl.length) return "Нет данных";
    var lastY=yl[yl.length-1];
    var arr=regionsWithData().map(function(r){ return {region:r,val:valueAt(r,lastY)||0}; }).sort(function(a,b){return b.val-a.val;}).slice(0,5);
    return arr.map(function(r){return r.region+": "+fmtNum(r.val,1);}).join("\n");
  }},
  stats:{cat:"Анализ",desc:"KPI всех регионов",usage:"stats",run:function(){
    return buildStats().map(function(r){ return r.region+": Σ="+fmtNum(r.sum,1)+", ср="+fmtNum(r.mean,2)+", CAGR="+(r.cagr!=null?fmtNum(r.cagr,2)+"%":"—"); }).join("\n")||"Нет данных";
  }},
  cagr:{cat:"Анализ",desc:"CAGR региона",usage:"cagr <регион>",run:function(a){ if(!a[0]) return "Укажите регион"; var s=seriesFor(a[0]); var c=cagrSeries(s.years,s.values); return c!=null?a[0]+": "+fmtNum(c,2)+"%/год":"Нет данных"; }},
  forecast:{cat:"Анализ",desc:"Прогноз",usage:"forecast <регион> [лет]",run:function(a){ if(!a[0]) return "Укажите регион"; var s=seriesFor(a[0]); if(!s.values.length) return "Нет данных"; var h=Number(a[1])||3, best=pickBestModel(s.values); if(!best) return "Мало данных"; var out="Прогноз "+a[0]+" ("+best.name+"):\n"; for(var k=1;k<=h;k++) out+="Y+"+k+": "+fmtNum(best.predict(s.values.length-1+k),2)+"\n"; return out; }},
  corr:{cat:"Анализ",desc:"Корреляция двух регионов",usage:"corr <r1> <r2>",run:function(a){ if(!a[0]||!a[1]) return "Укажите два"; var years=yearsAll().filter(function(y){return valueAt(a[0],y)!=null&&valueAt(a[1],y)!=null;}); if(years.length<2) return "Мало общих лет"; var r=pearson(years.map(function(y){return valueAt(a[0],y);}), years.map(function(y){return valueAt(a[1],y);})); return "r = "+(r!=null?fmtNum(r,3):"—"); }},
  gini:{cat:"Анализ",desc:"Джини",usage:"gini",run:function(){
    var yl=yearsAll(); if(!yl.length) return "Нет данных";
    var lastY=yl[yl.length-1], vals=REGIONS.filter(function(r){return state.selected[r];}).map(function(r){return valueAt(r,lastY);}).filter(function(v){return v!=null;});
    var g=gini(vals); return "Джини ("+lastY+"): "+(g!=null?fmtNum(g,3):"—");
  }},
  hhi:{cat:"Анализ",desc:"HHI",usage:"hhi [год]",run:function(a){
    var yl=yearsAll(); if(!yl.length) return "Нет данных";
    var y=Number(a[0])||yl[yl.length-1];
    var vals=REGIONS.filter(function(r){return state.selected[r];}).map(function(r){return valueAt(r,y);}).filter(function(v){return v!=null;});
    var tot=sum(vals); if(!tot) return "Нет данных";
    var h=0; vals.forEach(function(v){ var s=v/tot*100; h+=s*s; });
    return "HHI ("+y+"): "+fmtNum(h,0);
  }},
  mk:{cat:"Анализ",desc:"Манн-Кендалл",usage:"mk <регион>",run:function(a){ if(!a[0]) return "Укажите регион"; var s=seriesFor(a[0]), mk=mannKendall(s.values); if(mk.Z==null) return "Мало данных"; var tr=mk.Z>1.96?"рост":(mk.Z<-1.96?"падение":"нет"); return "S="+mk.S+"; Z="+fmtNum(mk.Z,2)+"; p="+fmtNum(mk.p,4)+"; тренд: "+tr; }},
  theme:{cat:"Система",desc:"Сменить тему",usage:"theme dark|light",run:function(a){ var t=a[0]||"dark"; state.theme=t; document.documentElement.setAttribute("data-theme",t); refreshAll(); return "Тема: "+t; }},
  select:{cat:"Система",desc:"Выбор регионов",usage:"select all|clear|only-with-data",run:function(a){
    var m=a[0]||"all";
    if(m==="all") REGIONS.forEach(function(r){state.selected[r]=true;});
    else if(m==="clear") state.selected={};
    else if(m==="only-with-data"){ state.selected={}; regionsWithData().forEach(function(r){state.selected[r]=true;}); }
    renderAllChips(); refreshAll();
    return "Выбрано: "+Object.keys(state.selected).filter(function(k){return state.selected[k];}).length;
  }},
  open:{cat:"Система",desc:"Перейти на страницу",usage:"open <page>",run:function(a){ if(!a[0]) return "Укажите страницу"; goTo(a[0]); return "Открыто: "+a[0]; }},
  docs:{cat:"Система",desc:"Руководство (docs.html)",usage:"docs",run:function(){ window.open("docs.html","_blank","noopener"); return "Открыто"; }},
  sources:{cat:"Система",desc:"Источники (sources.html)",usage:"sources",run:function(){ window.open("sources.html","_blank","noopener"); return "Открыто"; }},
  methodology:{cat:"Система",desc:"Методика (methodology.html)",usage:"methodology",run:function(){ window.open("methodology.html","_blank","noopener"); return "Открыто"; }},
  changelog:{cat:"Система",desc:"История (changelog.html)",usage:"changelog",run:function(){ window.open("changelog.html","_blank","noopener"); return "Открыто"; }},
  license:{cat:"Система",desc:"Лицензия (license.html)",usage:"license",run:function(){ window.open("license.html","_blank","noopener"); return "Открыто"; }},
  privacy:{cat:"Система",desc:"Политика (privacy.html)",usage:"privacy",run:function(){ window.open("privacy.html","_blank","noopener"); return "Открыто"; }},
  about:{cat:"Система",desc:"О программе",usage:"about",run:function(){ return "СППР «Солнечная Энергетика РФ» · v6.1 · MIT · 2026"; }},
  export:{cat:"Экспорт",desc:"Экспорт в формате",usage:"export <fmt>",run:function(a){ if(!a[0]) return "Форматы: png, jpeg, svg, pdf, html, csv, xlsx, json, docx, md, zip"; try{ var r=exportData(a[0]); if(r&&r.then) r.catch(function(e){tErr("Export: "+e.message);}); return "Экспорт "+a[0].toUpperCase()+" запущен"; }catch(e){ tErr(e.message); return "Ошибка: "+e.message; } }},
  save:{cat:"Экспорт",desc:"Скачать проект",usage:"save",run:function(){ var p={data:state.data, selected:state.selected, theme:state.theme, econ:state.econ}; downloadBlob(new Blob([JSON.stringify(p,null,2)],{type:"application/json"}), fileBase()+"_project.json"); return "Проект скачан"; }},
  undo:{cat:"Система",desc:"Отменить",usage:"undo",run:function(){ undo(); return "Отменено"; }},
  redo:{cat:"Система",desc:"Повторить",usage:"redo",run:function(){ redo(); return "Повторено"; }},
  "check libs":{cat:"Система",desc:"Библиотеки",usage:"check libs",run:function(){ return "Plotly:"+(plotlyReady()?"✓":"✗")+" XLSX:"+(xlsxReady()?"✓":"✗")+" docx:"+(docxReady()?"✓":"✗")+" JSZip:"+(jszipReady()?"✓":"✗"); }},
  reload:{cat:"Система",desc:"Перезагрузить",usage:"reload",run:function(){ location.reload(); return ""; }}
};
function runCommand(line){
  var parts=line.trim().split(/\s+/), cmd=parts[0], args=parts.slice(1);
  if(!cmd) return;
  var multi=["load demo","check libs"];
  for(var i=0;i<multi.length;i++) if(line.startsWith(multi[i])){
    tlog("CMD","$ "+line);
    try{ var o=COMMANDS[multi[i]].run(args); if(o) tlog("OK",o); }catch(e){ tErr(e.message); }
    return;
  }
  if(COMMANDS[cmd]){
    tlog("CMD","$ "+line);
    try{ var o2=COMMANDS[cmd].run(args); if(o2) tlog("OK",o2); }catch(e){ tErr(e.message); }
  } else { tlog("CMD","$ "+line); tErr("Неизвестная: "+cmd); }
}
function renderCommandCatalog(){
  var panel=$("panelCmd"); if(!panel) return;
  var groups={};
  Object.keys(COMMANDS).forEach(function(k){ var c=COMMANDS[k]; var cat=c.cat||"Прочее"; if(!groups[cat]) groups[cat]=[]; groups[cat].push({name:k,cmd:c}); });
  var order=["Данные","Анализ","Экспорт","Система","Основные","Прочее"];
  var html='<div class="t-muted" style="margin-bottom:10px">Всего: <b>'+Object.keys(COMMANDS).length+'</b>. Клик — вставить в поле.</div>';
  order.forEach(function(cat){
    if(!groups[cat]) return;
    html+='<div class="cmd-cat">'+escapeHtml(cat)+'</div><div class="cmd-grid">';
    groups[cat].forEach(function(item){
      html+='<div class="cmd-item" data-cmd="'+escapeHtml(item.name)+'"><div class="cmd-name">'+escapeHtml(item.name)+'</div><div class="cmd-desc">'+escapeHtml(item.cmd.desc)+'</div><div class="cmd-usage">'+escapeHtml(item.cmd.usage||item.name)+'</div></div>';
    });
    html+='</div>';
  });
  panel.innerHTML=html;
  qsa('#panelCmd .cmd-item').forEach(function(el){
    el.addEventListener("click",function(){ var cmd=el.getAttribute("data-cmd"); $("termInput").value=cmd; $("termInput").focus(); });
  });
}
function renderHelpPanel(){
  var p=$("panelHelp"); if(!p) return;
  p.innerHTML='<div class="t-muted"><b>Шпаргалка</b></div>'+
    '<div class="line"><span class="t-ok">~</span> — терминал</div>'+
    '<div class="line"><span class="t-ok">Ctrl+K</span> — палитра</div>'+
    '<div class="line"><span class="t-ok">Ctrl+Z/Y</span> — undo/redo</div>'+
    '<div class="line"><span class="t-ok">↑↓</span> — история команд</div>'+
    '<div class="line"><span class="t-ok">Tab</span> — автодополнение</div>'+
    '<div class="line"><span class="t-ok">Esc</span> — закрыть</div>';
}

/* ---------- DATA LOAD ---------- */
function loadDemo(){
  var demo=[], base={
    "Астраханская обл.":[180,220,260,290,320],
    "Волгоградская обл.":[210,250,290,330,370],
    "Оренбургская обл.":[280,330,380,410,440],
    "Самарская обл.":[140,170,200,230,260],
    "Саратовская обл.":[120,150,180,210,240],
    "Республика Башкортостан":[100,120,140,160,190],
    "Республика Алтай":[15,20,25,30,35],
    "Республика Крым":[300,340,380,400,420],
    "Ставропольский край":[260,290,330,360,400],
    "Краснодарский край":[220,250,280,310,340],
    "Ростовская обл.":[180,200,230,260,290],
    "Белгородская обл.":[60,80,100,120,140],
    "Республика Дагестан":[40,55,70,90,110],
    "Чеченская Республика":[30,45,60,80,100],
    "Республика Бурятия":[50,70,90,110,130],
    "Забайкальский край":[60,80,100,120,140],
    "Приморский край":[25,35,45,55,65]
  };
  var years=[2019,2020,2021,2022,2023];
  Object.keys(base).forEach(function(r){
    base[r].forEach(function(v,i){ demo.push({region:r, year:years[i], value:v, status:years[i]<2023?"оценка":"оценка", note:"требует проверки"}); });
  });
  state.data=demo;
  regionsWithData().forEach(function(r){ state.selected[r]=true; });
  renderAllChips(); refreshAll();
  tInfo("Демо загружено");
}
function loadSeed(){
  if(typeof window.SEED_DATA==="undefined"){ toast("data.js не найден","warn"); return; }
  var seed=window.SEED_DATA;
  if(!Array.isArray(seed)||!seed.length){ toast("data.js пуст","err"); return; }
  var res=validateRows(seed);
  if(!res.clean.length){ toast("Нет валидных строк в data.js","err"); return; }
  pushHistory(currentSnapshot());
  state.data=mergeData(state.data, res.clean);
  res.clean.forEach(function(r){ state.selected[r.region]=true; });
  renderAllChips(); refreshAll();
  toast("Загружено из базы: "+res.clean.length+" строк","ok");
}

/* ---------- PALETTE ---------- */
var PALETTE_ITEMS=[
  {cat:"Главная",name:"Открыть главную",action:function(){ goTo("home"); }},
  {cat:"Данные",name:"Загрузить базу СЭС (data.js)",action:function(){ loadSeed(); }},
  {cat:"Данные",name:"Демо-данные",action:function(){ pushHistory(currentSnapshot()); loadDemo(); }},
  {cat:"Данные",name:"Открыть ввод данных",action:function(){ goTo("data"); }},
  {cat:"Экономика",name:"Открыть окупаемость СЭС",action:function(){ goTo("economics"); }},
  {cat:"Экономика",name:"Сбросить параметры экономики",action:function(){ econFormReset(); showEconEmpty(); toast("Параметры экономики сброшены"); }},
  {cat:"Анализ",name:"Графики выработки",action:function(){ goTo("view"); }},
  {cat:"Анализ",name:"Тренды и прогнозы",action:function(){ goTo("trend"); }},
  {cat:"Тема",name:"Светлая",action:function(){ state.theme="light"; document.documentElement.setAttribute("data-theme","light"); refreshAll(); }},
  {cat:"Тема",name:"Тёмная",action:function(){ state.theme="dark"; document.documentElement.setAttribute("data-theme","dark"); refreshAll(); }},
  {cat:"Экспорт",name:"PDF",action:function(){ try{ exportData("pdf"); }catch(e){ tErr(e.message); } }},
  {cat:"Экспорт",name:"DOCX",action:function(){ try{ exportData("docx"); }catch(e){ tErr(e.message); } }},
  {cat:"Экспорт",name:"XLSX",action:function(){ try{ exportData("xlsx"); }catch(e){ tErr(e.message); } }},
  {cat:"Экспорт",name:"ZIP по регионам",action:function(){ try{ exportData("zip"); }catch(e){ tErr(e.message); } }},
  {cat:"Справка",name:"Руководство (docs.html)",action:function(){ window.open("docs.html","_blank","noopener"); }},
  {cat:"Справка",name:"Источники (sources.html)",action:function(){ window.open("sources.html","_blank","noopener"); }},
  {cat:"Справка",name:"Методика (methodology.html)",action:function(){ window.open("methodology.html","_blank","noopener"); }},
  {cat:"Справка",name:"История версий (changelog.html)",action:function(){ window.open("changelog.html","_blank","noopener"); }},
  {cat:"Справка",name:"Лицензия (license.html)",action:function(){ window.open("license.html","_blank","noopener"); }},
  {cat:"Справка",name:"Политика данных (privacy.html)",action:function(){ window.open("privacy.html","_blank","noopener"); }}
];
function openPalette(){ $("paletteBackdrop").classList.add("open"); $("palette").classList.add("open"); $("paletteInput").value=""; $("paletteInput").focus(); renderPaletteResults(""); }
function closePalette(){ $("paletteBackdrop").classList.remove("open"); $("palette").classList.remove("open"); }
function renderPaletteResults(q){
  q=(q||"").toLowerCase();
  var items=PALETTE_ITEMS.filter(function(it){ return !q||it.name.toLowerCase().indexOf(q)>=0||it.cat.toLowerCase().indexOf(q)>=0; });
  $("paletteResults").innerHTML=items.map(function(it,i){
    return '<div class="item'+(i===0?" selected":"")+'" data-idx="'+i+'"><span>'+escapeHtml(it.name)+'</span><span class="cat">'+escapeHtml(it.cat)+'</span></div>';
  }).join("")||'<div class="item">Ничего не найдено</div>';
  qsa("#paletteResults .item").forEach(function(el){
    el.addEventListener("click",function(){ var idx=+el.getAttribute("data-idx"); if(items[idx]){ try{items[idx].action();}catch(e){tErr(e.message);} } closePalette(); });
  });
}

/* ---------- DRAWER ---------- */
function openDrawer(){ $("drawerBackdrop").classList.add("open"); }
function closeDrawer(){ $("drawerBackdrop").classList.remove("open"); }

/* ---------- CONSENT ---------- */
function showConsent(){ $("consent").classList.add("open"); }
function hideConsent(){ $("consent").classList.remove("open"); }
function setConsent(kind){
  state.consent=kind;
  try{ localStorage.setItem(CONSENT_KEY, kind); }catch(e){}
  hideConsent();
  if(kind==="all"){ autoload(); loadSettingsToUI(); document.documentElement.setAttribute("data-theme",state.theme); refreshAll(); toast("Согласие принято. Автосохранение активно.","ok"); }
  else toast("Только необходимые. Автосохранение отключено.","ok");
}
function loadConsent(){
  try{ var c=localStorage.getItem(CONSENT_KEY); if(c==="all"||c==="necessary"){ state.consent=c; return true; } }catch(e){}
  return false;
}
function loadSettingsToUI(){
  if($("optUnit")) $("optUnit").value=state.optUnit;
  if($("optPrec")) $("optPrec").value=state.optPrec;
  if($("optSep")) $("optSep").value=state.optSep;
  if($("optTheme")) $("optTheme").value=state.theme;
}

/* ---------- SPLASH ---------- */
function setupSplashCanvas(){
  var canvas=$("splash-canvas"); if(!canvas) return;
  var ctx=canvas.getContext("2d"), W, H, particles=[];
  function resize(){ W=canvas.width=window.innerWidth; H=canvas.height=window.innerHeight; }
  resize(); window.addEventListener("resize", resize);
  for(var i=0;i<80;i++) particles.push({x:Math.random()*W,y:Math.random()*H,vx:(Math.random()-0.5)*0.35,vy:(Math.random()-0.5)*0.35,r:Math.random()*1.8+0.4});
  function draw(){
    ctx.clearRect(0,0,W,H);
    for(var i=0;i<particles.length;i++){
      var p=particles[i]; p.x+=p.vx; p.y+=p.vy;
      if(p.x<0||p.x>W) p.vx*=-1; if(p.y<0||p.y>H) p.vy*=-1;
      ctx.beginPath(); ctx.arc(p.x,p.y,p.r,0,Math.PI*2);
      ctx.fillStyle="rgba(255,255,255,"+(0.3+Math.random()*0.3)+")"; ctx.fill();
    }
    for(var a=0;a<particles.length;a++) for(var b=a+1;b<particles.length;b++){
      var dx=particles[a].x-particles[b].x, dy=particles[a].y-particles[b].y, d=Math.sqrt(dx*dx+dy*dy);
      if(d<100){ ctx.beginPath(); ctx.strokeStyle="rgba(253,224,71,"+(0.18*(1-d/100))+")"; ctx.lineWidth=0.5; ctx.moveTo(particles[a].x,particles[a].y); ctx.lineTo(particles[b].x,particles[b].y); ctx.stroke(); }
    }
    requestAnimationFrame(draw);
  }
  draw();
  var pctEl=$("splashPct"), start=Date.now(), dur=2200;
  function upd(){ var el=Date.now()-start, pct=Math.min(100,Math.round(el/dur*100)); if(pctEl) pctEl.textContent=pct+"%"; if(pct<100) requestAnimationFrame(upd); }
  requestAnimationFrame(upd);
}
function hideSplash(){ var s=$("splash"); if(!s) return; s.classList.add("hide"); setTimeout(function(){ if(s) s.remove(); }, 700); }

/* ---------- BIND ---------- */
function safe(fn,label){
  return function(ev){ try{ fn(ev); }catch(e){ tErr("["+label+"]: "+e.message); toast("Ошибка: "+e.message,"err"); } };
}
function bindAll(){
  on($("btnThemeLight"),"click",safe(function(){ state.theme="light"; document.documentElement.setAttribute("data-theme","light"); refreshAll(); },"tL"));
  on($("btnThemeDark"),"click",safe(function(){ state.theme="dark"; document.documentElement.setAttribute("data-theme","dark"); refreshAll(); },"tD"));
  on($("btnUndo"),"click",safe(undo,"undo"));
  on($("btnRedo"),"click",safe(redo,"redo"));

  qsa("[data-page]").forEach(function(a){
    on(a,"click",function(ev){ ev.preventDefault(); goTo(a.getAttribute("data-page")); });
  });
  on($("btnBack"),"click",safe(function(){ if(window.history.length>1) window.history.back(); else goTo("home"); },"back"));
  on($("mobileMoreBtn"),"click",openDrawer);
  on($("drawerClose"),"click",closeDrawer);
  on($("drawerBackdrop"),"click",function(e){ if(e.target===$("drawerBackdrop")) closeDrawer(); });
  on($("drawerThemeLight"),"click",safe(function(){ state.theme="light"; document.documentElement.setAttribute("data-theme","light"); refreshAll(); },"dTl"));
  on($("drawerThemeDark"),"click",safe(function(){ state.theme="dark"; document.documentElement.setAttribute("data-theme","dark"); refreshAll(); },"dTd"));

  /* Data page */
  on($("btnSelectAll"),"click",safe(function(){ REGIONS.forEach(function(r){state.selected[r]=true;}); renderAllChips(); refreshAll(); },"sa"));
  on($("btnClearSel"),"click",safe(function(){ state.selected={}; renderAllChips(); refreshAll(); },"sc"));
  on($("btnSelectWithData"),"click",safe(function(){ state.selected={}; regionsWithData().forEach(function(r){state.selected[r]=true;}); renderAllChips(); refreshAll(); },"sd"));
  on($("btnLoadSeed"),"click",safe(loadSeed,"seed"));
  on($("btnAddRow"),"click",safe(function(){ pushHistory(currentSnapshot()); state.data.push({region:REGIONS[0],year:new Date().getFullYear(),value:0,status:"факт",note:""}); refreshAll(); },"add"));
  on($("btnClearData"),"click",safe(function(){ if(!confirm("Удалить все данные?")) return; pushHistory(currentSnapshot()); state.data=[]; state.selected={}; renderAllChips(); refreshAll(); toast("Очищено"); },"clr"));
  on($("btnLoadDemo"),"click",safe(function(){ pushHistory(currentSnapshot()); loadDemo(); },"demo"));
  on($("btnPreviewPaste"),"click",safe(function(){
    try{
      var parsed=parseFlexible($("pasteArea").value);
      var res=validateRows(parsed);
      if(!res.clean.length){ toast("Нет валидных строк","err"); return; }
      if(!confirm("Добавить/обновить "+res.clean.length+" строк?\nОшибок: "+res.errors.length+"\nПредупреждений: "+res.warnings.length)) return;
      pushHistory(currentSnapshot());
      state.data=mergeData(state.data,res.clean);
      res.clean.forEach(function(r){ state.selected[r.region]=true; });
      renderAllChips(); refreshAll();
      toast("Добавлено: "+res.clean.length, res.errors.length?"warn":"ok");
      res.errors.forEach(tWarn); res.warnings.forEach(tWarn);
    }catch(e){ toast("Ошибка: "+e.message,"err"); tErr(e.message); }
  },"paste"));
  on($("btnPickFile"),"click",safe(function(){ $("fileInput").click(); },"pick"));
  on($("fileInput"),"change",safe(function(){
    var f=$("fileInput").files[0]; if(!f) return;
    var handle=function(parsed){
      var res=validateRows(parsed);
      if(!res.clean.length){ toast("Нет валидных","err"); return; }
      pushHistory(currentSnapshot());
      state.data=mergeData(state.data,res.clean);
      res.clean.forEach(function(r){ state.selected[r.region]=true; });
      renderAllChips(); refreshAll();
      toast("Из файла: "+res.clean.length);
    };
    if(/\.(xlsx|xls)$/i.test(f.name)){
      if(!xlsxReady()){ toast("XLSX недоступен","err"); return; }
      var fr=new FileReader();
      fr.onload=function(e){
        try{
          var wb=XLSX.read(new Uint8Array(e.target.result),{type:"array"});
          var ws=wb.Sheets[wb.SheetNames[0]];
          var rows=XLSX.utils.sheet_to_json(ws,{header:1,defval:""});
          var start=0; if(rows.length&&/регион|region/i.test(String(rows[0][0]||""))) start=1;
          handle(rows.slice(start).filter(function(r){return r.length>=3&&r[0]!=="";}).map(function(r){
            return {region:String(r[0]).trim(),year:Number(r[1]),value:Number(String(r[2]).replace(",",".")),status:String(r[3]||"факт"),note:String(r[4]||"")};
          }));
        }catch(err){ toast("Ошибка XLSX: "+err.message,"err"); }
      };
      fr.readAsArrayBuffer(f);
    } else {
      var fr2=new FileReader();
      fr2.onload=function(e){ try{ handle(parseFlexible(e.target.result)); }catch(err){ toast("Ошибка: "+err.message,"err"); } };
      fr2.readAsText(f);
    }
    $("fileInput").value="";
  },"file"));
  on($("btnTplCsv"),"click",safe(function(){
    var rows=[["Регион","Год","Значение","Статус","Примечание"]];
    REGIONS.slice(0,3).forEach(function(r){ [2021,2022,2023].forEach(function(y){ rows.push([r,y,"","факт",""]); }); });
    downloadBlob(new Blob(["\uFEFF"+rows.map(function(r){return r.join(";");}).join("\n")],{type:"text/csv;charset=utf-8"}),"шаблон_СЭС.csv");
  },"tplc"));
  on($("btnTplXlsx"),"click",safe(function(){
    if(!xlsxReady()){ toast("XLSX недоступен","err"); return; }
    var rows=[["Регион","Год","Значение","Статус","Примечание"]];
    REGIONS.forEach(function(r){ [2021,2022,2023].forEach(function(y){ rows.push([r,y,"","факт",""]); }); });
    var ws=XLSX.utils.aoa_to_sheet(rows); ws["!cols"]=[{wch:32},{wch:8},{wch:16},{wch:12},{wch:30}];
    var wb=XLSX.utils.book_new(); XLSX.utils.book_append_sheet(wb,ws,"Данные"); XLSX.writeFile(wb,"шаблон_СЭС.xlsx");
  },"tplx"));

  /* DeepSeek */
  if($("promptText")) $("promptText").textContent="Ты — старший аналитик ТЭК РФ. Извлеки из приложенных документов данные о годовой выработке солнечных электростанций по регионам РФ. Верни строгий JSON-массив:\n[{\"region\":\"Оренбургская обл.\",\"year\":2023,\"value\":410,\"status\":\"факт\",\"note\":\"СО ЕЭС\"}]\nЕдиница — млн кВт·ч. Регион — строго из официального названия. Только фактические данные. Округление до 0.1.";
  on($("btnOpenDS"),"click",safe(function(){ window.open("https://chat.deepseek.com/","_blank","noopener"); },"ds"));
  on($("btnCopyPrompt"),"click",safe(function(){
    var pr=$("promptText").textContent, done=function(){ toast("Скопировано","ok"); };
    var fb=function(){ var ta=document.createElement("textarea"); ta.value=pr; document.body.appendChild(ta); ta.select(); try{document.execCommand("copy"); done();}catch(e){toast("Ошибка","err");} ta.remove(); };
    if(navigator.clipboard&&navigator.clipboard.writeText) navigator.clipboard.writeText(pr).then(done,fb); else fb();
  },"copy"));

  /* View */
  on($("btnApplyView"),"click",safe(function(){ renderMainChart(); },"av"));
  on($("btnResetView"),"click",safe(function(){ if($("yearMin")) $("yearMin").value=""; if($("yearMax")) $("yearMax").value=""; if($("chartType")) $("chartType").value="lines"; if($("yScale")) $("yScale").value="linear"; if(plotlyReady()&&$("chart")) Plotly.relayout($("chart"),{xaxis:{autorange:true}, yaxis:{autorange:true}}); renderMainChart(); },"rv"));
  on($("chartType"),"change",safe(function(){ renderMainChart(); },"ct"));
  on($("yScale"),"change",safe(function(){ renderMainChart(); },"ys"));
  on($("yearMin"),"change",safe(function(){ renderMainChart(); },"ym"));
  on($("yearMax"),"change",safe(function(){ renderMainChart(); },"yM"));
  on($("detailRegion"),"change",safe(renderDetail,"det"));
  on($("btnBuildTrend"),"click",safe(function(){ renderTrend(); },"tbt"));
  on($("btnResetTrend"),"click",safe(function(){ if(plotlyReady()&&$("trendChart")) Plotly.relayout($("trendChart"),{xaxis:{autorange:true}, yaxis:{autorange:true}}); },"rt"));
  on($("forecastYears"),"change",safe(function(){ renderTrend(); },"fy"));
  on($("histRegion"),"change",safe(function(){ renderHeatmap(); },"hist"));
  on($("passRegion"),"change",safe(function(){ renderRegionPassport(); },"pass"));
  on($("btnBuildLag"),"click",safe(function(){ buildLagChart(); },"lag"));
  on($("btnBuildScenario"),"click",safe(function(){ renderScenario(); },"sc"));
  on($("scenRate"),"change",safe(function(){ renderScenario(); },"sr"));
  on($("scenYears"),"change",safe(function(){ renderScenario(); },"sy"));

  /* Economics */
  on($("btnEconApply"),"click",safe(calcEconomics,"econ"));
  on($("btnEconReset"),"click",safe(function(){
    econFormReset();
    showEconEmpty();
    state.econ = {};
    autosave();
    toast("Все поля сброшены к значениям по умолчанию","ok");
  },"econres"));
  if($("econWaccMode")) on($("econWaccMode"),"change",safe(function(){
    var waccFields=$("econWaccFields");
    if(waccFields) waccFields.style.display = $("econWaccMode").checked ? "grid" : "none";
    if($("econDiscount")) $("econDiscount").disabled = $("econWaccMode").checked;
  },"wacc"));

  /* Settings */
  on($("btnSaveSettings"),"click",safe(function(){
    state.optUnit=$("optUnit").value; state.optPrec=Number($("optPrec").value);
    state.optSep=$("optSep").value; state.theme=$("optTheme").value;
    document.documentElement.setAttribute("data-theme",state.theme);
    refreshAll(); toast("Настройки сохранены","ok");
  },"ss"));
  on($("btnResetSettings"),"click",safe(function(){
    state.optUnit="mln"; state.optPrec=2; state.optSep="ru"; state.theme="light";
    loadSettingsToUI(); document.documentElement.setAttribute("data-theme","light"); refreshAll();
  },"rs"));
  on($("btnOpenConsent"),"click",showConsent);
  on($("consentAccept"),"click",function(){ setConsent("all"); });
  on($("consentNecessary"),"click",function(){ setConsent("necessary"); });

  /* Project */
  on($("btnSaveProject"),"click",safe(function(){
    var p={data:state.data, selected:state.selected, theme:state.theme, optUnit:state.optUnit, optPrec:state.optPrec, optSep:state.optSep, econ:state.econ};
    downloadBlob(new Blob([JSON.stringify(p,null,2)],{type:"application/json"}), fileBase()+"_project.json");
  },"sp"));
  on($("btnLoadProject"),"click",safe(function(){ $("projFile").click(); },"lp"));
  on($("projFile"),"change",safe(function(){
    var f=$("projFile").files[0]; if(!f) return;
    var fr=new FileReader();
    fr.onload=function(e){
      try{
        var p=JSON.parse(e.target.result);
        pushHistory(currentSnapshot());
        if(p.data) state.data=p.data;
        if(p.selected) state.selected=p.selected;
        if(p.theme) state.theme=p.theme;
        if(p.optUnit) state.optUnit=p.optUnit;
        if(p.optPrec) state.optPrec=p.optPrec;
        if(p.optSep) state.optSep=p.optSep;
        if(p.econ) state.econ=p.econ;
        document.documentElement.setAttribute("data-theme",state.theme);
        loadSettingsToUI(); renderAllChips(); refreshAll();
        toast("Проект загружен");
      }catch(err){ toast("Ошибка: "+err.message,"err"); }
    };
    fr.readAsText(f); $("projFile").value="";
  },"lpj"));

  /* Export */
  qsa('input[name="expMode"]').forEach(function(r){
    on(r,"change",safe(function(){ state.expMode=r.value; },"em"));
  });
  qsa("[data-exp]").forEach(function(btn){
    on(btn,"click",safe(function(){
      var fmt=btn.getAttribute("data-exp");
      try{
        var res=exportData(fmt);
        if(res&&res.then) res.catch(function(e){ toast("Ошибка "+fmt+": "+e.message,"err"); });
      }catch(e){ toast("Ошибка "+fmt+": "+e.message,"err"); }
    },"exp-"+btn.getAttribute("data-exp")));
  });

  /* Caption */
  ["figCaption","figSource","sigAuthor","sigOrg","sigDate","sigNote","sigCity"].forEach(function(id){
    var el=$(id); if(!el) return;
    on(el,"input",safe(updateCaption,"cap"));
    on(el,"change",safe(updateCaption,"cap"));
  });

  /* Terminal */
  on($("termToggle"),"click",function(){ $("terminal").classList.toggle("open"); });
  on($("termClose"),"click",function(){ $("terminal").classList.remove("open"); });
  qsa(".term-tabs button[data-tab]").forEach(function(b){
    on(b,"click",function(){
      var tab=b.getAttribute("data-tab");
      qsa(".term-tabs button[data-tab]").forEach(function(x){ x.classList.toggle("active",x===b); });
      qsa(".term-tab-panel").forEach(function(p){ p.classList.toggle("active",p.getAttribute("data-panel")===tab); });
      if(tab==="cmd") renderCommandCatalog();
      if(tab==="help") renderHelpPanel();
    });
  });
  qsa(".term-quick button").forEach(function(b){ on(b,"click",function(){ runCommand(b.getAttribute("data-cmd")); }); });
  var ti=$("termInput"), hist=[], hi=-1;
  if(ti){
    on(ti,"keydown",function(e){
      if(e.key==="Enter"){ var v=ti.value.trim(); if(v){ hist.push(v); hi=hist.length; runCommand(v); } ti.value=""; }
      else if(e.key==="ArrowUp"){ if(hi>0){hi--; ti.value=hist[hi]||"";} e.preventDefault(); }
      else if(e.key==="ArrowDown"){ if(hi<hist.length-1){hi++; ti.value=hist[hi]||"";} else{hi=hist.length;ti.value="";} e.preventDefault(); }
      else if(e.key==="Tab"){
        e.preventDefault(); var v=ti.value; if(!v) return;
        var pre=v.split(/\s+/)[0], keys=Object.keys(COMMANDS), m=keys.filter(function(k){return k.indexOf(pre)===0;})[0];
        if(m) ti.value=m+v.slice(pre.length);
      }
    });
  }
  renderCommandCatalog(); renderHelpPanel();

  /* Palette */
  on($("paletteBackdrop"),"click",closePalette);
  on($("paletteInput"),"input",function(e){ renderPaletteResults(e.target.value); });
  on($("paletteInput"),"keydown",function(e){
    if(e.key==="Escape") closePalette();
    if(e.key==="Enter"){ var s=document.querySelector("#paletteResults .item.selected"); if(s) s.click(); }
  });

  /* Keyboard */
  document.addEventListener("keydown",function(e){
    if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==="k"){ e.preventDefault(); openPalette(); return; }
    if((e.ctrlKey||e.metaKey)&&e.key.toLowerCase()==="z"&&!e.shiftKey){
      if(document.activeElement.tagName!=="INPUT"&&document.activeElement.tagName!=="TEXTAREA"){ e.preventDefault(); undo(); }
    }
    if((e.ctrlKey||e.metaKey)&&(e.key.toLowerCase()==="y"||(e.shiftKey&&e.key.toLowerCase()==="z"))){
      if(document.activeElement.tagName!=="INPUT"&&document.activeElement.tagName!=="TEXTAREA"){ e.preventDefault(); redo(); }
    }
    if(e.key==="`"&&!e.ctrlKey&&!e.metaKey&&document.activeElement.tagName!=="INPUT"&&document.activeElement.tagName!=="TEXTAREA"){ e.preventDefault(); $("terminal").classList.toggle("open"); }
    if(e.key==="Escape"){ closePalette(); closeDrawer(); }
  });

  on($("onbSkip"),"click",closeOnb);
  on($("onbStart"),"click",function(){ closeOnb(); });
  on($("onbBackdrop"),"click",closeOnb);

  window.addEventListener("popstate",function(ev){
    var p = (ev.state&&ev.state.page) || (location.hash||"").replace(/^#/,"") || "home";
    if(p&&p!=="home") goTo(p); else goTo("home");
  });
}
function closeOnb(){ $("onb").classList.remove("open"); $("onbBackdrop").classList.remove("open"); try{ localStorage.setItem("solar_rf_onb","1"); }catch(e){} }

function updateCaption(){
  var cap=$("figCaption")?$("figCaption").value:"";
  var src=$("figSource")?$("figSource").value:"";
  if($("figureCaption")) $("figureCaption").textContent=cap;
  if($("figureSource")) $("figureSource").textContent=src;
  if($("figureCaption2")) $("figureCaption2").textContent=cap;
  if($("figureSource2")) $("figureSource2").textContent=src;
  if($("sigAuthorOut")) $("sigAuthorOut").textContent=($("sigAuthor")&&$("sigAuthor").value)||"\u00A0";
  if($("sigOrgOut")) $("sigOrgOut").textContent=($("sigOrg")&&$("sigOrg").value)||"\u00A0";
  if($("sigDateOut")){ var d=$("sigDate")?$("sigDate").value:""; $("sigDateOut").textContent=d?fmtDate(d):"\u00A0"; }
}

/* ---------- START ---------- */
function checkLibs(){
  var miss=[];
  if(!plotlyReady()) miss.push("Plotly");
  if(!xlsxReady()) miss.push("SheetJS");
  if(!docxReady()) miss.push("docx");
  if(miss.length){
    $("banner").className="show";
    $("banner").innerHTML="Не загружены: <code>"+miss.join("</code>, <code>")+"</code>. Расчёты и CSV/JSON/HTML работают.";
    tWarn("Не загружены: "+miss.join(", "));
  } else $("banner").className="";
}
function start(){
  setupSplashCanvas();
  var skip=$("spSplashSkip"); var timer=null;
  if(skip) skip.addEventListener("click",function(){ clearTimeout(timer); hideSplash(); });
  timer=setTimeout(hideSplash, 2400);

  checkLibs();
  bindAll();

  renderAllChips();
  updateCaption();
  populateRegionSelects();

  var consentLoaded=loadConsent();
  if(consentLoaded){
    autoload();
    loadSettingsToUI();
    document.documentElement.setAttribute("data-theme",state.theme);
    refreshAll();
    var initialPage = state.currentPage || (location.hash||"").replace(/^#/,"") || "home";
    goTo(initialPage);
  } else {
    setTimeout(function(){ showConsent(); refreshAll(); goTo("home"); }, 2500);
  }

  tInfo("СППР «Солнечная Энергетика РФ» запущена · v6.1");
  setTimeout(function(){
    if(!localStorage.getItem("solar_rf_onb") && consentLoaded){
      $("onb").classList.add("open"); $("onbBackdrop").classList.add("open");
    }
  }, 3200);

  /* Restore econ from saved state */
  if(state.econ && state.econ.power){
    setTimeout(function(){
      if($("econPower")) $("econPower").value = state.econ.power;
      if($("econYield")) $("econYield").value = state.econ.yieldPerKW;
      if($("econCapex")) $("econCapex").value = state.econ.capex;
      if($("econOpex")) $("econOpex").value = state.econ.opexPct;
      if($("econTariff")) $("econTariff").value = state.econ.tariff;
      if($("econDegr")) $("econDegr").value = state.econ.degr*100;
      if($("econDiscount")) $("econDiscount").value = state.econ.discount*100;
      if($("econLife")) $("econLife").value = state.econ.life;
      if($("econTariffGrowth")) $("econTariffGrowth").value = state.econ.tariffGrowth*100;
      if($("econOpexGrowth")) $("econOpexGrowth").value = state.econ.opexGrowth*100;
      if($("econTax")) $("econTax").value = state.econ.tax*100;
      if($("econSubsidy")) $("econSubsidy").value = state.econ.subsidy;
      if($("econDegrType")) $("econDegrType").value = state.econ.degrType;
      if($("econRegion") && state.econ.region) $("econRegion").value = state.econ.region;
      if($("econWaccMode")) $("econWaccMode").checked = state.econ.waccMode;
      if($("econDebtPct")) $("econDebtPct").value = state.econ.debtPct*100;
      if($("econDebtRate")) $("econDebtRate").value = state.econ.debtRate*100;
      if($("econEquityRate")) $("econEquityRate").value = state.econ.equityRate*100;
      if($("econWaccFields")) $("econWaccFields").style.display = state.econ.waccMode ? "grid" : "none";
      if($("econDiscount")) $("econDiscount").disabled = !!state.econ.waccMode;
    }, 200);
  }
}
if(document.readyState==="loading") document.addEventListener("DOMContentLoaded", start);
else start();

})();
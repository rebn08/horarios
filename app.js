/* Generador de horarios FIME · 100% client-side */
const $=s=>document.querySelector(s),KEY='fime_horarios_v1';
const DAYS=['LUNES','MARTES','MIÉRCOLES','JUEVES','VIERNES'];
const ROWS={M:[['07:00-7:50'],['07:50-8:40'],['08:40-9:30'],['9:30-10:00',1],['10:00-10:50'],['10:50-11:40'],['11:40-12:30'],['12:30-13:20'],['13:20-14:10']],
 V:[['4:10 - 5:00'],['5:00-5:50'],['5:50-6:40'],['6:40-7:30'],['7:30-8:20'],['8:20-9:10'],['9:10-10:00']]};
const TIMES={M:[[420,470],[470,520],[520,570],0,[600,650],[650,700],[700,750],[750,800],[800,850]],V:[[970,1020],[1020,1070],[1070,1120],[1120,1170],[1170,1220],[1220,1270],[1270,1320]]};
const ROWIDX={M:[0,1,2,4,5,6,7,8],V:[0,1,2,3,4,5,6]};
const PAL=['#ff99ff','#66ffff','#b4a7d6','#ffe599','#ffff00','#99ff66','#c5d9a0','#f9bf8f','#8db4e2','#e0bde8','#92d050','#3399ff','#ffb3b3','#b6f0d0'];
const norm=s=>String(s??'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/[^a-z0-9]/g,'');
const esc=s=>String(s??'').replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
const tn=v=>/^v/i.test(v)?'V':'M';
const fresh=()=>({cfg:{periodo:'AGO-DIC 2026',dir:'DR. JESUS ROLANDO RODRIGUEZ SILVA',sec:'DR. JESUS AARON RIOS GALINDO',logoL:'',logoR:''},salones:[],docentes:[],materias:[],horarios:{},warn:[]});
let S=Object.assign(fresh(),(()=>{try{return JSON.parse(localStorage[KEY])}catch(e){return{}}})());
const save=()=>{try{localStorage[KEY]=JSON.stringify(S)}catch(e){toast('Almacenamiento lleno: exporta tu proyecto (.json)',1)}};
function toast(m,bad){const t=$('#toast');t.textContent=m;t.className=bad?'bad':'';t.hidden=false;clearTimeout(toast.t);toast.t=setTimeout(()=>t.hidden=true,3800)}
let tab='datos',sub='salones',cur='',sel=null,editing=false,GM={};

/* ---------- Datos de ejemplo / plantillas ---------- */
const DN=['GARCIA ALVARADO ADRIAN','GONZALEZ VALDEZ LETICIA','BURCIAGA ALARCON RICARDO','CHAVANA GUZMAN NANCY LIZETH','MARTINEZ ZUÑIGA RODOLFO','RODRIGUEZ GAUNA JORGE RAUL','MACGREGOR CAMPUZANO MARCELA','SERVIN CASTAÑEDA RUMUALDO','ALMA LUCIA GARCIA MATA'];
const mk=(c,s,sc,t,al,mat,d,ht,hl)=>({carrera:c,sem:s,sec:sc,turno:t,alumnos:al,materia:mat,docente:d,ht,hl});
const SAMPLE={salones:[{nombre:'29',tipo:'Aula',cap:40},{nombre:'7',tipo:'Aula',cap:40},{nombre:'LAB-1',tipo:'Laboratorio',cap:30}],
 docentes:DN.map((n,i)=>({id:'E'+(1001+i),nombre:n,d0:'07:00-22:00',d1:'07:00-22:00',d2:'07:00-22:00',d3:'07:00-22:00',d4:'07:00-22:00'})),
 materias:[mk('IB',1,'A','Matutino',35,'Circuitos Logicos y Secuenciales',DN[0],3,2),mk('IB',1,'A','Matutino',35,'Ingles Pre-Inicial',DN[1],3,0),mk('IB',1,'A','Matutino',35,'Fundamentos de Programacion',DN[2],3,0),mk('IB',1,'A','Matutino',35,'Fisica',DN[3],3,2),mk('IB',1,'A','Matutino',35,'Fundamentos de Matematicas',DN[4],5,0),
 mk('IME',5,'C','Vespertino',30,'Ingenieria Termica',DN[5],3,0),mk('IME',5,'C','Vespertino',30,'Ingles V',DN[6],3,0),mk('IME',5,'C','Vespertino',30,'Mecanica de Materiales II',DN[7],3,0),mk('IME',5,'C','Vespertino',30,'Ecuaciones Diferenciales',DN[8],5,0)]};
function template(){const wb=XLSX.utils.book_new(),ad=(n,a)=>XLSX.utils.book_append_sheet(wb,XLSX.utils.aoa_to_sheet(a),n);
 ad('Salones',[['Nombre','Tipo','Capacidad'],...SAMPLE.salones.map(s=>[s.nombre,s.tipo,s.cap])]);
 ad('Docentes',[['Expediente','Nombre',...DAYS],...SAMPLE.docentes.map(d=>[d.id,d.nombre,d.d0,d.d1,d.d2,d.d3,d.d4])]);
 ad('Materias',[['Carrera','Semestre','Seccion','Turno','Alumnos','Materia','Docente','Horas Teoria','Horas Lab'],...SAMPLE.materias.map(m=>[m.carrera,m.sem,m.sec,m.turno,m.alumnos,m.materia,m.docente,m.ht,m.hl])]);
 ad('Instrucciones',[['Disponibilidad docente: escribe "HH:MM-HH:MM" (ej. 07:00-14:10; varios rangos separados por coma). Vacío = sin restricción. "X" = no disponible ese día.'],['Tipo de salón: Aula o Laboratorio. Turno: Matutino o Vespertino. Horas Lab > 0 = requiere laboratorio.']]);
 XLSX.writeFile(wb,'plantillas_horarios_FIME.xlsx')}
function impXlsx(wb){const sh=n=>{const k=wb.SheetNames.find(x=>norm(x).startsWith(n));return k?XLSX.utils.sheet_to_json(wb.Sheets[k],{defval:''}).map(r=>{const o={};for(const c in r)o[norm(c)]=r[c];return o}):null};
 let a=sh('salon');if(a)S.salones=a.filter(r=>r.nombre!=='').map(r=>({nombre:String(r.nombre),tipo:/lab/i.test(r.tipo)?'Laboratorio':'Aula',cap:+r.capacidad||0}));
 a=sh('docente');if(a)S.docentes=a.filter(r=>r.nombre).map(r=>({id:String(r.expediente||r.id||''),nombre:String(r.nombre),d0:r.lunes,d1:r.martes,d2:r.miercoles,d3:r.jueves,d4:r.viernes}));
 a=sh('materia');if(a)S.materias=a.filter(r=>r.materia).map(r=>({carrera:String(r.carrera),sem:+r.semestre||1,sec:String(r.seccion),turno:/^v/i.test(r.turno)?'Vespertino':'Matutino',alumnos:+r.alumnos||0,materia:String(r.materia),docente:String(r.docente),ht:+r.horasteoria||0,hl:+r.horaslab||0}));
 S.horarios={};save()}

/* ---------- Algoritmo generador ---------- */
const findDoc=s=>{const n=norm(s);return S.docentes.find(d=>norm(d.id)==n||norm(d.nombre)==n)};
function okAvail(doc,d,t,b){if(!doc)return true;const r=String(doc['d'+d]||'').trim();if(!r)return true;if(/^(x|-|no)/i.test(r))return false;
 const [a,e]=TIMES[t][ROWIDX[t][b]];return r.split(',').some(p=>{const m=p.match(/(\d{1,2}):(\d{2})\s*-\s*(\d{1,2}):(\d{2})/);return m&&a>=+m[1]*60+ +m[2]&&e<=+m[3]*60+ +m[4]})}
function groupsOf(){const m={};GM={};S.materias.forEach(x=>{const t=tn(x.turno),k=[x.carrera,x.sem,x.sec,t].join('|');const g=m[k]=m[k]||{k,t,carrera:x.carrera,sem:x.sem,sec:x.sec,alumnos:0,mats:[],label:`${x.sem}º ${x.carrera} "${x.sec}"`};g.alumnos=Math.max(g.alumnos,+x.alumnos||0);g.mats.push(x)});
 const l=Object.values(m).sort((a,b)=>(a.carrera+a.t).localeCompare(b.carrera+b.t)||a.sem-b.sem||String(a.sec).localeCompare(b.sec));l.forEach(g=>GM[g.k]=g);return l}
function attempt(gs,it){const bT={},bR={},h={},un=[],use={};let cost=0;
 const order=gs.map(g=>[g,g.mats.reduce((a,m)=>a+(+m.ht||0)+(+m.hl||0),0)+Math.random()*it*6]).sort((a,b)=>b[1]-a[1]).map(x=>x[0]);
 for(const g of order){const t=g.t,R=ROWIDX[t],nb=R.length;h[g.k]=[];
  const occ=[0,1,2,3,4].map(()=>Array(nb).fill(false));
  const aulas=S.salones.filter(s=>s.tipo!='Laboratorio'&&+s.cap>=g.alumnos).sort((a,b)=>(use[t+a.nombre]||0)-(use[t+b.nombre]||0)||a.cap-b.cap);
  const labs=S.salones.filter(s=>s.tipo=='Laboratorio'&&+s.cap>=g.alumnos).sort((a,b)=>a.cap-b.cap);
  const hm=aulas[0];if(hm)use[t+hm.nombre]=(use[t+hm.nombre]||0)+1;
  const tasks=[];g.mats.forEach(m=>{const doc=findDoc(m.docente);[[+m.ht||0,0],[+m.hl||0,1]].forEach(([n,lab])=>{while(n>0){const l=n>=2?2:1;tasks.push({m,doc,lab,l});n-=l}})});
  tasks.sort((a,b)=>b.l-a.l||b.lab-a.lab||Math.random()-.5);
  const place=(k,len)=>{let best=null;const dk=norm(k.m.docente);const pool=k.lab?(labs.length?labs:aulas):[hm,...aulas];
   for(let d=0;d<5;d++)for(let b=0;b+len<=nb;b++){let ok=true;
    for(let i=0;i<len&&ok;i++){const bb=b+i;if(occ[d][bb]||(i&&R[bb]!=R[bb-1]+1)||bT[`${t}|${dk}|${d}|${bb}`]||!okAvail(k.doc,d,t,bb))ok=false}
    if(!ok)continue;
    const room=pool.find(r=>r&&[...Array(len)].every((_,i)=>!bR[`${t}|${r.nombre}|${d}|${b+i}`]));if(!room)continue;
    const tmp=occ[d].slice();for(let i=0;i<len;i++)tmp[b+i]=true;const f=tmp.indexOf(true),l=tmp.lastIndexOf(true);let gaps=0;for(let i=f;i<=l;i++)if(!tmp[i])gaps++;
    const same=h[g.k].filter(c=>c.mat==k.m.materia&&c.d==d);let pen=0;if(same.length)pen=same.some(c=>c.b+1==b||b+len==c.b)?12:30;
    const sc=gaps*10+f*1.2+occ[d].filter(Boolean).length*.8+pen+Math.random()*3;
    if(!best||sc<best.sc)best={sc,d,b,room}}
   if(!best)return false;
   for(let i=0;i<len;i++){occ[best.d][best.b+i]=true;bT[`${t}|${dk}|${best.d}|${best.b+i}`]=1;bR[`${t}|${best.room.nombre}|${best.d}|${best.b+i}`]=1;
    h[g.k].push({d:best.d,b:best.b+i,mat:k.m.materia,doc:k.doc?k.doc.nombre:k.m.docente,lab:k.lab,room:best.room.nombre})}
   return true};
  for(const k of tasks){let ok=place(k,k.l);if(!ok&&k.l==2)ok=place(k,1)&&place(k,1);if(!ok){un.push(`${g.label}: «${k.m.materia}» (${k.lab?'laboratorio':'teoría'}, ${k.l} h) sin ubicar: revisa disponibilidad del docente, salones o capacidad`);cost+=1000}}
  occ.forEach(o=>{const f=o.indexOf(true),l=o.lastIndexOf(true);for(let i=f;i>=0&&i<=l;i++)if(!o[i])cost+=10})}
 return{h,un,cost}}
function generate(){const gs=groupsOf();let best=null;
 for(let it=0;it<30;it++){const r=attempt(gs,it);if(!best||r.un.length<best.un.length||(r.un.length==best.un.length&&r.cost<best.cost))best=r;if(!best.un.length&&it>=10)break}
 S.horarios=best?best.h:{};S.warn=best?best.un:[];save()}

/* ---------- Render del horario oficial ---------- */
function sheet(g,merge,edit){const t=g.t,R=ROWS[t],cells=S.horarios[g.k]||[],n=R.length;
 const at=[0,1,2,3,4].map(()=>Array(n).fill(null)),span=[0,1,2,3,4].map(()=>Array(n).fill(1)),cs=[0,1,2,3,4].map(()=>Array(n).fill(1)),hid=[0,1,2,3,4].map(()=>Array(n).fill(false));
 cells.forEach(c=>at[c.d][ROWIDX[t][c.b]]=c);const key=c=>c?[c.mat,c.doc,c.lab,c.room].join('|'):'';
 if(merge){for(let d=0;d<5;d++)for(let r=0;r<n;r++)if(at[d][r]&&!hid[d][r]){let k=1;while(r+k<n&&at[d][r+k]&&key(at[d][r+k])==key(at[d][r]))k++;span[d][r]=k;for(let j=1;j<k;j++)hid[d][r+j]=true}
  for(let r=0;r<n;r++){if([0,1,2,3,4].every(d=>at[d][r]&&!hid[d][r]&&span[d][r]==1&&key(at[d][r])==key(at[0][r]))){cs[0][r]=5;for(let d=1;d<5;d++)hid[d][r]=true}}}
 const names=[...new Set(g.mats.map(m=>m.materia))],col=m=>PAL[names.indexOf(m)%PAL.length];
 const home={};cells.filter(c=>!c.lab).forEach(c=>home[c.room]=(home[c.room]||0)+1);const salon=Object.keys(home).sort((a,b)=>home[b]-home[a])[0]||(cells[0]||{}).room||'';
 const rh=Math.floor((520-(t=='M'?38:0))/(t=='M'?8:7));
 const rows=R.map((row,r)=>{if(row[1])return`<tr style="height:38px"><td class="h">${row[0]}</td>${DAYS.map(()=>'<td>RECESO</td>').join('')}</tr>`;
  return`<tr style="height:${rh}px"><td class="h">${row[0]}</td>${[0,1,2,3,4].map(d=>{if(hid[d][r])return'';const c=at[d][r];
   const pos=edit?` data-d="${d}" data-r="${r}"`:'';const cls=edit?`c ${sel&&sel.d==d&&sel.r==r?'sel':''}`:'';
   if(!c)return`<td class="${cls}"${pos}></td>`;
   return`<td class="${cls}"${pos} ${edit?'draggable="true"':''} rowspan="${span[d][r]}" colspan="${cs[d][r]}" style="background:${col(c.mat)}"><div>${esc(c.mat)}</div><div class="d">${esc(c.doc)}</div>${c.lab?'<div class="lb">LABORATORIO · '+esc(c.room)+'</div>':''}</td>`}).join('')}</tr>`}).join('');
 const L=S.cfg.logoL?`<img src="${S.cfg.logoL}">`:'<div class="lg"></div>',Rr=S.cfg.logoR?`<img src="${S.cfg.logoR}">`:'<div class="lg"></div>';
 return`<div class="sheet"><div class="hd">${L}<div class="tt"><b>UNIVERSIDAD AUTÓNOMA DE COAHUILA</b><span>FACULTAD DE INGENIERÍA MECÁNICA Y ELÉCTRICA</span></div>${Rr}</div>
 <div class="chips"><div>GRUPO:<span class="chip">${esc(g.label)}</span></div><div>SALÓN:<span class="chip">${esc(salon)}</span></div><div><span class="chip">${esc(S.cfg.periodo)}</span></div></div>
 <table><tr><th style="width:92px"></th>${DAYS.map(d=>`<th>${d}</th>`).join('')}</tr>${rows}</table>
 <div class="sig"><div>${esc(S.cfg.dir)}<small>DIRECTOR</small></div><div>${esc(S.cfg.sec)}<small>SECRETARIO ACADEMICO</small></div></div></div>`}

/* ---------- Edición manual con validación ---------- */
function conflict(c,d,b,g){if(!okAvail(findDoc(c.doc),d,g.t,b))return`${c.doc} no está disponible a esa hora`;
 for(const gk in S.horarios){if(gk==g.k||!GM[gk]||GM[gk].t!=g.t)continue;for(const o of S.horarios[gk])if(o.d==d&&o.b==b){
  if(norm(o.doc)==norm(c.doc))return`Choque: ${c.doc} ya imparte a ${GM[gk].label} a esa hora`;if(o.room==c.room)return`Choque: el salón ${c.room} está ocupado por ${GM[gk].label}`}}return''}
function tryMove(f,to,apply){const g=GM[cur],cl=S.horarios[cur],a=cl.find(c=>c.d==f.d&&c.b==f.b);if(!a)return'';
 const bc=cl.find(c=>c.d==to.d&&c.b==to.b);if(a===bc)return'';
 const e=conflict(a,to.d,to.b,g)||(bc?conflict(bc,f.d,f.b,g):'');if(e||!apply)return e;
 a.d=to.d;a.b=to.b;if(bc){bc.d=f.d;bc.b=f.b}save();return''}
const pos=td=>({d:+td.dataset.d,b:ROWIDX[GM[cur].t].indexOf(+td.dataset.r),r:+td.dataset.r});
function cellClick(td){const p=pos(td);const has=S.horarios[cur].some(c=>c.d==p.d&&c.b==p.b);
 if(!sel){if(has)sel=p}else{if(sel.d==p.d&&sel.b==p.b)sel=null;else{const e=tryMove(sel,p,1);sel=null;if(e)toast(e,1)}}render()}
document.addEventListener('dragstart',e=>{const td=e.target.closest&&e.target.closest('td.c');if(td)sel=pos(td)});
document.addEventListener('dragover',e=>{if(e.target.closest&&e.target.closest('td.c'))e.preventDefault()});
document.addEventListener('dragenter',e=>{const td=e.target.closest&&e.target.closest('td.c');if(td&&sel)td.classList.add(tryMove(sel,pos(td))?'bad':'ok')});
document.addEventListener('dragleave',e=>{const td=e.target.closest&&e.target.closest('td.c');if(td)td.classList.remove('ok','bad')});
document.addEventListener('drop',e=>{const td=e.target.closest&&e.target.closest('td.c');if(!td||!sel)return;e.preventDefault();const er=tryMove(sel,pos(td),1);sel=null;if(er)toast(er,1);render()});

/* ---------- Exportación ---------- */
const dl=(b,n)=>{const a=document.createElement('a');a.href=URL.createObjectURL(b);a.download=n;a.click();setTimeout(()=>URL.revokeObjectURL(a.href),5e3)};
const fn=g=>'Horario_'+g.label.replace(/[^\wÁ-ú]+/g,'_')+'_'+g.t;
const cap=g=>{const c=$('#cap');c.innerHTML=sheet(g,true,false);return html2canvas(c.firstElementChild,{scale:2,backgroundColor:'#fff'})};
async function makePdf(list){const p=new jspdf.jsPDF({orientation:'landscape',unit:'in',format:'letter'});for(let i=0;i<list.length;i++){const c=await cap(list[i]);if(i)p.addPage();p.addImage(c.toDataURL('image/jpeg',.95),'JPEG',0,0,11,8.5)}return p}
const withSched=()=>groupsOf().filter(g=>S.horarios[g.k]);

/* ---------- UI ---------- */
const btn=(a,t,c='bg-blue-600',x='')=>`<button data-act="${a}" ${x} class="${c} text-white rounded-xl px-4 py-2 font-semibold">${t}</button>`;
const COLS={salones:[['nombre','Salón / Lab'],['tipo','Tipo',['Aula','Laboratorio']],['cap','Butacas','number']],
 docentes:[['id','Expediente'],['nombre','Nombre completo'],...DAYS.map((D,i)=>['d'+i,D.slice(0,3)+' (HH:MM-HH:MM)'])],
 materias:[['carrera','Carrera'],['sem','Sem.','number'],['sec','Sección'],['turno','Turno',['Matutino','Vespertino']],['alumnos','Alumnos','number'],['materia','Materia'],['docente','Docente (nombre o exp.)'],['ht','Hrs teoría','number'],['hl','Hrs lab','number']]};
const BLANK={salones:{nombre:'',tipo:'Aula',cap:30},docentes:{id:'',nombre:'',d0:'',d1:'',d2:'',d3:'',d4:''},materias:{carrera:'',sem:1,sec:'A',turno:'Matutino',alumnos:30,materia:'',docente:'',ht:3,hl:0}};
function renderDatos(){const cols=COLS[sub];
 return`<div class="flex gap-2 mb-3 flex-wrap">${['salones','docentes','materias'].map(k=>btn('sub',{salones:'🏫 Salones y laboratorios',docentes:'👩‍🏫 Docentes',materias:'📚 Materias y grupos'}[k],k==sub?'bg-blue-700':'bg-slate-500',`data-k="${k}"`)).join('')}</div>
 <div class="overflow-auto bg-white rounded-xl p-2 shadow"><table class="w-full text-sm"><thead><tr>${cols.map(c=>`<th class="p-1 text-left">${c[1]}</th>`).join('')}<th></th></tr></thead><tbody>
 ${S[sub].map((r,i)=>`<tr>${cols.map(([f,,o])=>`<td class="p-1">${Array.isArray(o)?`<select class="inp" data-k="${sub}" data-i="${i}" data-f="${f}" data-n="0">${o.map(v=>`<option ${v==r[f]?'selected':''}>${v}</option>`).join('')}</select>`:`<input class="inp" ${o?'type="number" data-n="1"':'data-n="0"'} value="${esc(r[f])}" data-k="${sub}" data-i="${i}" data-f="${f}">`}</td>`).join('')}<td><button data-act="del" data-k="${sub}" data-i="${i}" class="text-red-600 text-xl px-2">✕</button></td></tr>`).join('')}</tbody></table>
 <div class="mt-3">${btn('add','＋ Agregar fila','bg-green-600',`data-k="${sub}"`)}</div></div>
 ${sub=='docentes'?'<p class="mt-2 text-sm">Disponibilidad: «07:00-14:10» (varios rangos con coma). Vacío = sin restricción. «X» = no disponible.</p>':''}`}
function renderGen(){const gs=groupsOf();return`<div class="bg-white rounded-xl p-4 shadow space-y-3"><p class="text-lg">${gs.length} grupos · ${S.materias.length} materias · ${S.docentes.length} docentes · ${S.salones.length} salones</p>
 ${btn('gen','⚙ Generar horarios automáticamente','bg-green-600')}
 ${S.warn.length?`<div class="bg-red-50 border border-red-300 rounded-xl p-3"><b>Pendientes (${S.warn.length}):</b><ul class="list-disc ml-5">${S.warn.map(w=>`<li>${esc(w)}</li>`).join('')}</ul></div>`:withSched().length?'<p class="text-green-700 font-semibold">✔ Todas las clases fueron ubicadas. Ve a «Horarios».</p>':''}</div>`}
function renderHor(){const gs=withSched();if(!gs.length)return'<p class="p-6 text-lg">Aún no hay horarios. Ve a «Generar».</p>';if(!GM[cur]||!S.horarios[cur])cur=gs[0].k;
 return`<div class="flex flex-wrap gap-2 items-center mb-3"><select id="grp" class="inp !w-auto text-lg">${gs.map(g=>`<option value="${esc(g.k)}" ${g.k==cur?'selected':''}>${esc(g.label)} · ${g.t=='M'?'Matutino':'Vespertino'}</option>`).join('')}</select>
 ${btn('edit',editing?'✔ Terminar edición':'✋ Editar manualmente',editing?'bg-amber-600':'bg-slate-600')}${btn('pdf','PDF')}${btn('jpg','JPG')}${btn('pdfAll','Todos en un PDF','bg-indigo-600')}${btn('zipAll','Todos en ZIP (JPG)','bg-indigo-600')}</div>
 ${editing?'<p class="mb-2">Toca una clase y luego la hora destino (o arrástrala). Si hay choque de docente, disponibilidad o salón, se bloquea y se avisa.</p>':''}
 <div class="overflow-auto bg-slate-300 p-2 rounded-xl"><div id="view" style="width:1100px;zoom:${Math.min(1,(innerWidth-40)/1100)}">${sheet(GM[cur],!editing,editing)}</div></div>`}
function renderCfg(){const c=S.cfg,f=(k,l)=>`<label class="block">${l}<input class="inp" data-cfg="${k}" value="${esc(c[k])}"></label>`;
 return`<div class="bg-white rounded-xl p-4 shadow grid gap-3 max-w-xl">${f('periodo','Periodo')}${f('dir','Firma izquierda (Director)')}${f('sec','Firma derecha (Secretario Académico)')}
 <label>Logo UAdeC (izquierda)<input type="file" accept="image/*" data-logo="logoL" class="block"></label><label>Logo FIME (derecha)<input type="file" accept="image/*" data-logo="logoR" class="block"></label></div>`}
const TABS={datos:['📋 Datos',renderDatos],gen:['⚙ Generar',renderGen],hor:['🗓 Horarios',renderHor],cfg:['🎨 Configuración',renderCfg]};
function render(){$('#nav').innerHTML=Object.entries(TABS).map(([k,[l]])=>`<button data-act="tab" data-k="${k}" class="px-4 py-2 rounded-t-xl font-semibold whitespace-nowrap ${k==tab?'bg-slate-100 text-blue-900':'bg-white/20'}">${l}</button>`).join('');$('#main').innerHTML=TABS[tab][1]()}
document.addEventListener('click',async e=>{const td=e.target.closest('td.c');if(td&&editing)return cellClick(td);
 const el=e.target.closest('[data-act]');if(!el)return;const a=el.dataset.act,k=el.dataset.k,i=+el.dataset.i;let rr=1;
 try{if(a=='tab')tab=k;else if(a=='sub')sub=k;else if(a=='add'){S[k].push({...BLANK[k]});save()}else if(a=='del'){S[k].splice(i,1);save()}
 else if(a=='gen'){toast('Generando…');await new Promise(r=>setTimeout(r,40));generate();toast(S.warn.length?'Generado con pendientes':'¡Horarios generados!',S.warn.length)}
 else if(a=='demo'){if(confirm('Se reemplazarán tus datos por el ejemplo. ¿Continuar?')){Object.assign(S,structuredClone(SAMPLE),{horarios:{},warn:[]});save()}}
 else if(a=='reset'){if(confirm('¿Borrar TODO el proyecto?')){S=fresh();save()}}
 else if(a=='tpl'){template();rr=0}else if(a=='expP'){dl(new Blob([JSON.stringify(S)],{type:'application/json'}),'proyecto_horarios_FIME.json');rr=0}
 else if(a=='impP'){$('#fP').click();rr=0}else if(a=='impX'){$('#fX').click();rr=0}
 else if(a=='edit'){editing=!editing;sel=null}
 else{rr=0;toast('Generando archivo…');
  if(a=='pdf')(await makePdf([GM[cur]])).save(fn(GM[cur])+'.pdf');
  else if(a=='jpg'){const c=await cap(GM[cur]);c.toBlob(b=>dl(b,fn(GM[cur])+'.jpg'),'image/jpeg',.95)}
  else if(a=='pdfAll')(await makePdf(withSched())).save('Horarios_FIME_todos.pdf');
  else if(a=='zipAll'){const z=new JSZip();for(const g of withSched()){const c=await cap(g);z.file(fn(g)+'.jpg',c.toDataURL('image/jpeg',.95).split(',')[1],{base64:true})}dl(await z.generateAsync({type:'blob'}),'Horarios_FIME_JPG.zip')}
  toast('Listo ✔')}}catch(err){toast('Error: '+err.message,1)}
 if(rr)render()});
document.addEventListener('change',e=>{const t=e.target,d=t.dataset;
 if(d.k&&d.f){S[d.k][+d.i][d.f]=d.n=='1'?+t.value:t.value;S.horarios={};save()}
 else if(d.cfg){S.cfg[d.cfg]=t.value;save()}
 else if(d.logo&&t.files[0]){const r=new FileReader();r.onload=()=>{S.cfg[d.logo]=r.result;save();toast('Logo cargado')};r.readAsDataURL(t.files[0])}
 else if(t.id=='grp'){cur=t.value;sel=null;render()}
 else if(t.id=='fP'&&t.files[0]){const r=new FileReader();r.onload=()=>{try{S=Object.assign(fresh(),JSON.parse(r.result));save();render();toast('Proyecto importado')}catch(x){toast('Archivo inválido',1)}};r.readAsText(t.files[0]);t.value=''}
 else if(t.id=='fX'&&t.files[0]){const r=new FileReader();r.onload=()=>{try{impXlsx(XLSX.read(r.result,{type:'array'}));render();toast('Excel cargado')}catch(x){toast('Error leyendo Excel: '+x.message,1)}};r.readAsArrayBuffer(t.files[0]);t.value=''}});
render();

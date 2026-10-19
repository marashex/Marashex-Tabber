const $=id=>document.getElementById(id),esc=s=>String(s).replace(/[&<>"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
let S={};let activeTab='';const closed=new Set();
async function load(){S=await chrome.storage.local.get(null);render();}
function render(){
  $('notifyOpen').checked=!!S.notifyOpen;$('notifyAny').checked=!!S.notifyAny;
  if(document.activeElement.id!=='sheetUrl')$('sheetUrl').value=S.sheetUrl||'';
  if(document.activeElement.id!=='icsUrl')$('icsUrl').value=S.icsUrl||'';
  if(document.activeElement.id!=='durMin')$('durMin').value=S.durMin||60;
  const st=S.status||'',lc=S.lastChecked?new Date(S.lastChecked).toLocaleTimeString():'never';
  $('status').textContent=st==='ok'?`Watching · checked ${lc}`:st==='signin'?'Sign in to Olin SharePoint in Chrome, then Check now':st?`Problem: ${st}`:'Starting…';
  $('ignoreAD').checked=S.ignoreAllDay!==false;if(!$('diag').open)renderDiag();
  $('redir').value=chrome.identity.getRedirectURL();
  if(document.activeElement.id!=='clientId')$('clientId').value=S.clientId||'';
  $('calchip').textContent=S.msConnected?(S.calStatus&&S.calStatus!=='ok'?'Outlook: '+S.calStatus:'Outlook connected ✓'):(S.icsUrl||(S.events||[]).length?'Calendar: ICS':'Connect Outlook');
  const star=S.starred||{},ev=S.events||[],dur=S.durMin||60;
  let tr=[...(S.trainings||[])];
  if($('onlyOpen').checked)tr=tr.filter(t=>t.open);if($('onlyStar').checked)tr=tr.filter(t=>star[t.key]);
  tr.sort((a,b)=>(!!star[b.key]-!!star[a.key])||(a.start||1e15)-(b.start||1e15));
  const meta=S.meta||{},tabs=Object.keys(meta),all=S.trainings||[];
  if(!tabs.includes(activeTab))activeTab=tabs.find(n=>all.some(t=>t.name===n&&t.open))||tabs[0]||'';
  const pills=tabs.map(n=>{const o=all.filter(t=>t.name===n&&t.open).length;return `<button class="tabp ${n===activeTab?'on':''}" data-tab="${esc(n)}">${esc(n)}<span>${meta[n]&&meta[n].skipped?'?':o}</span></button>`;}).join('');
  const mine=tr.filter(t=>t.name===activeTab),m=meta[activeTab]||{},allMine=all.filter(t=>t.name===activeTab);
  const mat=(allMine.find(t=>t.mat)||{}).mat||m.tabMat||'';
  $('warn').innerHTML=m.skipped?`<div class="warn">No dated sign-up times found on <b>${esc(activeTab)}</b> <button id="reviewTabs" data-n="${esc(activeTab)}">Review</button></div>`:'';
  const tile=t=>{
    const c=conflictsFor(t,ev,dur,S.ignoreAllDay!==false),n=parseFloat(t.avail),num=/^\s*\d/.test(t.avail)&&!isNaN(n);
    const cal=c===null?'<span class="when">no time to check</span>':!(S.msConnected||S.icsUrl||ev.length)?'<span class="when">connect calendar</span>':c.length?`<span class="busy" title="${esc(c.map(x=>x.t).join(', '))}">⚠ Conflict: ${esc(c[0].t)}</span>`:'<span class="free">✓ You\'re free</span>';
    return `<div class="tile ${t.open?'t-open':'t-full'} ${star[t.key]?'star':''}"><div class="ttype">${esc(t.type||t.name)}</div>${t.who?`<div class="who">with ${esc(t.who)}</div>`:''}<div class="tdate">${esc(t.date)} · ${esc(t.time)}</div>
      <div class="spots">${num?n:(t.open?'Open':'Full')}</div><div class="when">${num?(n===1?'spot open':'spots open'):(esc(t.avail)||(t.open?'spots available':'no spots'))}</div>
      <div class="trow">${cal}<button class="star-btn ${star[t.key]?'on':''}" data-k="${esc(t.key)}" title="Priority">${star[t.key]?'★':'☆'}</button></div></div>`;};
  const descs=(m.secs||[]).map(s=>`<div class="desc"><div class="ttype">${esc(s.type||'Overview')}</div><div class="when">${(s.info||[]).map(x=>typeof x==='string'?esc(x):(x.u?`<a class="lnk" href="${esc(x.u)}" target="_blank" rel="noopener">${esc(x.t)} ↗</a>`:esc(x.t))).join('<br>')}</div>${s.mat?`<a class="mat" href="${esc(s.mat)}" target="_blank" rel="noopener">Needed materials ↗</a>`:''}</div>`).join('');
  $('list').innerHTML=tabs.length?`<div class="tabs">${pills}</div><div class="tabbody"><div class="thead"><span class="gname">${esc(activeTab)}</span>${mat?`<a class="mat" href="${esc(mat)}" target="_blank" rel="noopener">Needed materials ↗</a>`:''}<span class="gsum">${allMine.filter(t=>t.open).length} of ${allMine.length} upcoming times open</span></div>
    <div class="tiles">${mine.map(tile).join('')||'<p class="when">No upcoming times to show here.</p>'}</div>
    <h3 class="dh">About these trainings</h3><div class="descs">${descs||'<p class="when">No description found on this tab.</p>'}</div></div>`:'<p class="when">No tabs read yet. Hit “Check now”.</p>';
}
$('list').onclick=async e=>{const p=e.target.closest('.tabp');if(p){activeTab=p.dataset.tab;render();return;}const b=e.target.closest('.star-btn');if(!b)return;e.preventDefault();const s={...(S.starred||{})};
  if(b.dataset.g){const ts=(S.trainings||[]).filter(t=>t.name===b.dataset.g),all=ts.every(t=>s[t.key]);ts.forEach(t=>all?delete s[t.key]:s[t.key]=1);}
  else{const k=b.dataset.k;s[k]?delete s[k]:s[k]=1;}
  await chrome.storage.local.set({starred:s});};
$('list').addEventListener('toggle',e=>{const n=e.target.dataset.n;if(n!==undefined)e.target.open?closed.delete(n):closed.add(n);},true);
for(const id of['notifyOpen','notifyAny'])$(id).onchange=e=>chrome.storage.local.set({[id]:e.target.checked});
for(const id of['onlyOpen','onlyStar'])$(id).onchange=render;
$('check').onclick=async()=>{$('status').textContent='Checking…';await chrome.runtime.sendMessage('check');};
$('openSheet').onclick=()=>chrome.tabs.create({url:S.sheetUrl});
$('save').onclick=async()=>{await chrome.storage.local.set({sheetUrl:$('sheetUrl').value.trim(),icsUrl:$('icsUrl').value.trim(),durMin:+$('durMin').value||60,hash:null});await chrome.runtime.sendMessage('check');};
$('icsFile').onchange=async e=>{const f=e.target.files[0];if(f)await chrome.storage.local.set({events:parseICS(await f.text()),icsUrl:''});};
chrome.storage.onChanged.addListener(load);load();

$('calchip').onclick=()=>{const d=$('settings');d.open=true;d.scrollIntoView({behavior:'smooth'});};
$('connect').onclick=async()=>{
  const id=$('clientId').value.trim();if(!id){$('calmsg').textContent='Paste the client ID first.';return;}
  await chrome.storage.local.set({clientId:id});$('calmsg').textContent='Opening Microsoft sign-in…';
  const r=await chrome.runtime.sendMessage('connect');$('calmsg').textContent=r&&r.ok?'Connected ✓':'Failed: '+(r&&r.err||'unknown');
};
$('disconnect').onclick=async()=>{await chrome.runtime.sendMessage('disconnect');$('calmsg').textContent='Disconnected.';};

let diagSheet='';
function renderDiag(){
  const pv=Array.isArray(S.preview)?{}:(S.preview||{}),meta=S.meta||{},names=Object.keys(pv),tr=S.trainings||[];
  if(!names.includes(diagSheet))diagSheet=names[0]||'';
  $('cSheet').innerHTML=names.map(n=>`<option ${n===diagSheet?'selected':''}>${esc(n)}</option>`).join('');
  const m=meta[diagSheet]||{},hs=m.headers||[],cfg=(S.cfg||{})[diagSheet]||{};
  const perTab=names.map(n=>{const x=meta[n]||{};return `${n}: ${x.dated||0} dated rows (${x.past||0} past)`;}).join(' | ');
  $('diagsum').textContent=perTab+' — '+`${names.length} tabs · ${tr.length} trainings read (${tr.filter(t=>t.open).length} open) · ${(S.events||[]).length} calendar events loaded`+(m.skipped?` · this tab looks like info only (no date or availability column found) — pick columns below if it’s a sign-up tab`:'');
  $('prev').innerHTML=(pv[diagSheet]||[]).map((r,i)=>`<tr class="${i===m.hi?'hd':''}">${r.map(x=>`<td>${esc(x)}</td>`).join('')}</tr>`).join('');
  const opts=sel=>`<option value="">(auto)</option><option value="-1">(none)</option>`+hs.map((h,i)=>`<option value="${i}" ${String(sel)===String(i)?'selected':''}>${i+1}: ${esc(h||'(blank)')}</option>`).join('');
  for(const[id,k]of[['cName','name'],['cDate','date'],['cTime','time'],['cAvail','avail']])$(id).innerHTML=opts(cfg[k]);
  $('cRule').value=cfg.rule||'auto';$('cCap').value=cfg.cap||'';
}
$('diag').addEventListener('toggle',()=>{if($('diag').open)renderDiag();});
$('cSheet').onchange=e=>{diagSheet=e.target.value;renderDiag();};
$('ignoreAD').onchange=e=>chrome.storage.local.set({ignoreAllDay:e.target.checked});
$('savecfg').onclick=async()=>{
  const all={...(S.cfg&&!S.cfg.rule&&!('name'in S.cfg)?S.cfg:{})};
  all[diagSheet]={name:$('cName').value,date:$('cDate').value,time:$('cTime').value,avail:$('cAvail').value,rule:$('cRule').value,cap:$('cCap').value};
  await chrome.storage.local.set({cfg:all,trainings:null,hash:null});await chrome.runtime.sendMessage('check');renderDiag();
};

$('warn').onclick=e=>{const b=e.target.closest('#reviewTabs');if(!b)return;diagSheet=b.dataset.n;const d=$('diag');d.open=true;renderDiag();d.scrollIntoView({behavior:'smooth'});};
$('copyLayout').onclick=async()=>{await navigator.clipboard.writeText(JSON.stringify(S.preview||{}));$('copyLayout').textContent='Copied ✓';setTimeout(()=>$('copyLayout').textContent='Copy tab layouts',2000);};

async function grabSheet(){
  const u=new URL(S.sheetUrl);u.searchParams.set('download','1');
  const r=await fetch(u.href,{credentials:'include'}),buf=await r.arrayBuffer();
  if(new Uint8Array(buf.slice(0,2)).join()!=='80,75')throw new Error('Not signed in to SharePoint in this Chrome profile');
  return buf;
}
$('dlSheet').onclick=async()=>{try{const buf=await grabSheet(),a=document.createElement('a');a.href=URL.createObjectURL(new Blob([buf]));a.download='shop-training-sheet.xlsx';a.click();$('dbgmsg').textContent='Saved to Downloads ✓';}catch(e){$('dbgmsg').textContent=e.message;}};
$('copyAll').onclick=async()=>{try{
  const book=await parseWorkbook(await grabSheet());
  const txt=book.map(b=>`=== TAB: ${b.name} ===\n`+b.rows.map((r,i)=>`${i+1}\t`+r.join('\t')).join('\n')+'\n'+Object.entries(b.rows.links||{}).map(([k,v])=>`LINK r${+k.split(',')[0]+1}c${+k.split(',')[1]+1}: ${v}`).join('\n')).join('\n\n');
  await navigator.clipboard.writeText(txt);$('dbgmsg').textContent=`Copied ${txt.length} characters ✓`;}catch(e){$('dbgmsg').textContent=e.message;}};

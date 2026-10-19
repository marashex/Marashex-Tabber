importScripts('lib.js');
const DEFAULT_URL='https://olincollege.sharepoint.com/:x:/s/OlinShopTraining/IQDf9LLiHHqgR4JiiC_QMfiBAdj0zKo83EK_hW6NapdCNag?e=PsGzwp&CID=d2d56dd9-7902-035f-faa0-af235bc10ac4';
const get=k=>chrome.storage.local.get(k),set=o=>chrome.storage.local.set(o);
const ensureAlarm=()=>chrome.alarms.create('poll',{periodInMinutes:1});
chrome.runtime.onInstalled.addListener(async()=>{
  const s=await get(['sheetUrl']);
  if(!s.sheetUrl)await set({sheetUrl:DEFAULT_URL,notifyOpen:true,notifyAny:false,starred:{},durMin:60});
  ensureAlarm();check();
});
chrome.runtime.onStartup.addListener(ensureAlarm);
chrome.alarms.onAlarm.addListener(a=>a.name==='poll'&&check());
chrome.action.onClicked.addListener(()=>chrome.tabs.create({url:chrome.runtime.getURL('tabber.html')}));
chrome.notifications.onClicked.addListener(()=>chrome.tabs.create({url:chrome.runtime.getURL('tabber.html')}));
chrome.runtime.onMessage.addListener((m,_s,res)=>{
  if(m==='check'){check().then(()=>res(true));return true;}
  if(m==='connect'){(async()=>{try{const s=await get(null);await getToken(s,true);await set({msConnected:true,calStatus:'ok',msFailAt:null});await check();res({ok:true});}catch(e){res({ok:false,err:e.message});}})();return true;}
  if(m==='disconnect'){set({msConnected:false,msTok:null,calStatus:null,events:[]}).then(()=>res(true));return true;}
});
async function getToken(s,interactive){
  const st=await get(['msTok','msExp']);
  if(!interactive&&st.msTok&&st.msExp>Date.now()+60000)return st.msTok;
  const p=new URLSearchParams({client_id:s.clientId,response_type:'token',redirect_uri:chrome.identity.getRedirectURL(),scope:'https://graph.microsoft.com/Calendars.Read',response_mode:'fragment'});
  if(!interactive)p.set('prompt','none');
  const r=await chrome.identity.launchWebAuthFlow({url:'https://login.microsoftonline.com/organizations/oauth2/v2.0/authorize?'+p,interactive});
  const h=new URLSearchParams(new URL(r).hash.slice(1));
  if(!h.get('access_token'))throw new Error(h.get('error_description')||'sign-in failed');
  await set({msTok:h.get('access_token'),msExp:Date.now()+(+h.get('expires_in')||3600)*1000});
  return h.get('access_token');
}
async function graphEvents(s){
  const tok=await getToken(s,false),out=[];
  let url='https://graph.microsoft.com/v1.0/me/calendarView?'+new URLSearchParams({startDateTime:new Date().toISOString(),endDateTime:new Date(Date.now()+60*864e5).toISOString(),'$top':'200','$select':'subject,start,end,showAs,isCancelled'});
  while(url&&out.length<2000){
    const r=await fetch(url,{headers:{Authorization:'Bearer '+tok,Prefer:'outlook.timezone="UTC"'}});
    if(r.status===401){await set({msTok:null});throw new Error('token expired');}
    const j=await r.json();if(j.error)throw new Error(j.error.message);
    for(const e of j.value)if(!e.isCancelled&&e.showAs!=='free')out.push({s:Date.parse(e.start.dateTime.replace(/\.\d+$/,'')+'Z'),e:Date.parse(e.end.dateTime.replace(/\.\d+$/,'')+'Z'),t:e.subject||'busy'});
    url=j['@odata.nextLink'];
  }
  return out;
}
function notify(title,message,priority=false){
  chrome.notifications.create('',{type:'basic',iconUrl:'icon128.png',title,message,priority:2,requireInteraction:priority});
}
async function check(){
  const s=await get(null),url=s.sheetUrl||DEFAULT_URL;
  try{
    const u=new URL(url);u.searchParams.set('download','1');
    const r=await fetch(u.href,{credentials:'include'});
    const buf=await r.arrayBuffer();
    if(new Uint8Array(buf.slice(0,2)).join()!=='80,75'){
      await set({status:'signin',lastChecked:Date.now()});
      if(!s.signinWarned){notify('maroshex tabber','Sign in to Olin SharePoint in Chrome so I can read the sheet.');await set({signinWarned:true});}
      return;
    }
    const book=await parseWorkbook(buf),tr=[],metas={},prevs={},js=JSON.stringify(book.map(b=>[b.rows,b.rows.links]));
    for(const sh of book){const t=toTrainings(sh.rows,(s.cfg||{})[sh.name]||{},sh.name);tr.push(...t);metas[sh.name]=t.meta;prevs[sh.name]=sh.rows.slice(0,15).map(r=>r.slice(0,14));}
    const hash=js.length+':'+[...js].reduce((h,c)=>(h*31+c.charCodeAt(0))|0,7);
    // calendar refresh (published Outlook ICS link)
    let events=s.events||[];
    if(s.clientId&&s.msConnected){
      if(!(s.msFailAt&&Date.now()-s.msFailAt<6e5)){
        try{events=await graphEvents(s);await set({calStatus:'ok',msFailAt:null});}
        catch(e){await set({calStatus:'needs reconnect ('+e.message+')',msFailAt:Date.now()});}
      }
    }else if(s.icsUrl){try{const t=await (await fetch(s.icsUrl)).text();events=parseICS(t);}catch(e){}}
    const first=!s.trainings,prev=Object.fromEntries((s.trainings||[]).map(t=>[t.key,t])),star=s.starred||{},dur=s.durMin||60;
    if(!first&&hash!==s.hash){
      const lines=[];
      for(const t of tr){
        const p=prev[t.key],conf=(conflictsFor(t,events,dur,s.ignoreAllDay!==false)||[]).length?' ⚠ conflicts with your calendar':'';
        const becameOpen=t.open&&(!p||!p.open);
        if(becameOpen&&(star[t.key]||s.notifyOpen)){
          const isStar=!!star[t.key];
          notify((isStar?'★ Priority: ':'Open: ')+(t.label||t.name),`${t.date} ${t.time} · ${t.avail||'open'}${conf}`,isStar);
        }
        if(p&&p.avail!==t.avail)lines.push(`${(t.label||t.name)}: ${p.avail||'–'} → ${t.avail||'–'}`);
        else if(!p)lines.push(`New: ${(t.label||t.name)}`);
      }
      for(const k in prev)if(!tr.find(t=>t.key===k))lines.push(`Removed: ${(prev[k].label||prev[k].name)}`);
      if(s.notifyAny)notify('Training sheet updated',lines.slice(0,4).join('\n')||'Sheet contents changed');
    }
    await set({meta:metas,preview:prevs,trainings:tr,hash,events,status:'ok',lastChecked:Date.now(),signinWarned:false,lastChange:hash!==s.hash?Date.now():s.lastChange});
  }catch(e){await set({status:'error:'+e.message,lastChecked:Date.now()});}
}

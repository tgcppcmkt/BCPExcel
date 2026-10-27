'use strict';
/* Operation-per-file sharing, following Workload Monitor's folder approach.
 * Unique append-only files avoid lost writes. Deterministic replay converges;
 * it is NOT a distributed mutex and late events can change an earlier result.
 */
const DeskFolder=(()=>{
 const APP='MentionRequestDesk',DIR='MENTION_DESK_DATA',PREFIX='mention-folder-pending-v2-',LEGACY_PREFIX='mention-folder-pending-v1-',ID=/^h-[a-f0-9]{32}$/,FILE=/^op-h-[a-f0-9]{32}\.json$/;
 const clone=x=>JSON.parse(JSON.stringify(x)),own=(o,k)=>Object.prototype.hasOwnProperty.call(o,k),cmp=(a,b)=>a<b?-1:a>b?1:0;
 const error=(message,code='storage',ids=[])=>Object.assign(new Error(message),{code,ids});
 const uid=()=>{const b=new Uint8Array(16);crypto.getRandomValues(b);return'h-'+Array.from(b,x=>x.toString(16).padStart(2,'0')).join('');};
 const token=s=>typeof s==='string'&&(s===''||ID.test(s));
 function dayKey(date=new Date()){return String(date.getFullYear()).padStart(4,'0')+String(date.getMonth()+1).padStart(2,'0')+String(date.getDate()).padStart(2,'0');}
 function validDay(day){if(typeof day!=='string'||!/^\d{8}$/.test(day))return false;const y=Number(day.slice(0,4)),m=Number(day.slice(4,6)),d=Number(day.slice(6,8)),date=new Date(y,m-1,d);return y>=1000&&date.getFullYear()===y&&date.getMonth()===m-1&&date.getDate()===d;}
 function validateAction(a){
  if(!a||typeof a!=='object'||!['import','start','complete','note'].includes(a.kind))throw error('共有履歴の操作が不正です。','integrity');
  if(a.kind==='import'){
   if(!Array.isArray(a.records)||a.records.length>500)throw error('共有履歴の取込データが不正です。','integrity');
   for(const r of a.records)if(!r||DeskCore.fields.some(f=>typeof r[f]!=='string'||r[f].length>30000)||!r.id.trim()||r.id!==r.id.trim()||r.id.length>250)throw error('共有履歴の案件項目が不正です。','integrity');
  }else if(a.kind==='note'){
   if(typeof a.id!=='string'||!a.id||typeof a.note!=='string'||a.note.length>10000||!Number.isSafeInteger(a.expectedRevision)||a.expectedRevision<0||!token(a.expectedToken))throw error('共有履歴のメモ情報が不正です。','integrity');
  }else{
   if(!Array.isArray(a.ids)||!a.ids.length||a.ids.length>100000||a.ids.some(id=>typeof id!=='string'||!id)||new Set(a.ids).size!==a.ids.length||!a.expected||!a.expectedTokens||typeof a.expected!=='object'||typeof a.expectedTokens!=='object'||a.ids.some(id=>!own(a.expected,id)||!Number.isSafeInteger(a.expected[id])||a.expected[id]<0||!own(a.expectedTokens,id)||!token(a.expectedTokens[id])))throw error('共有履歴の選択案件が不正です。','integrity');
   if(a.kind==='start'&&(typeof a.name!=='string'||!a.name.trim()||a.name.trim().length>80))throw error('共有履歴の担当者名が不正です。','integrity');
  }return a;
 }
 function validateEvent(e,name){
  if(!e||e.schemaVersion!==1||e.application!==APP||!ID.test(e.id)||name!=='op-'+e.id+'.json'||!Number.isSafeInteger(e.clock)||e.clock<1||e.clock>1000000000000||typeof e.stamp!=='string'||!/^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d\.\d{3}Z$/.test(e.stamp)||!Number.isFinite(Date.parse(e.stamp)))throw error('共有履歴の形式が不正です：'+name,'integrity');
  validateAction(e.action);return e;
 }
 function apply(db,event){
  const op=event.action;
  if(op.kind==='start'||op.kind==='complete'){
   const map=new Map(db.requests.map(r=>[r.id,r]));
   const bad=op.ids.filter(id=>!map.has(id)||(map.get(id).stateToken||'')!==op.expectedTokens[id]);
   if(bad.length)throw new DeskCore.Conflict(op.kind==='start'?'こちらの案件は対応が開始されています。案件を再選択してください。':'こちらの案件の担当・対応状況が変更されています。案件を再選択してください。',bad);
  }else if(op.kind==='note'){
   const r=db.requests.find(r=>r.id===op.id);
   if(!r||(r.noteToken||'')!==op.expectedToken)throw new DeskCore.Conflict('メモが別の担当者によって更新されました。最新のメモを確認してください。',[op.id]);
  }
  const ids=op.kind==='import'?new Set(db.requests.map(r=>r.id)):null;
  const result=DeskCore.transact(db,op,event.stamp);
  // Token links bind a completion to the exact claim it was based on.
  // A late competing claim with the same numeric revision cannot be completed.
  if(result.db===db)return result;
  if(op.kind==='import'){for(const r of result.db.requests)if(!ids.has(r.id)){r.originToken=event.id;r.stateToken='';r.noteToken='';}}
  else if(op.kind==='note'){result.db.requests.find(r=>r.id===op.id).noteToken=event.id;}
  else{const selected=new Set(op.ids);for(const r of result.db.requests)if(selected.has(r.id))r.stateToken=event.id;}
  return result;
 }
 function replay(events,previous=null){
  const ordered=events.slice().sort((a,b)=>a.clock-b.clock||cmp(a.id,b.id)),seen=new Set();
  for(const e of ordered){validateEvent(e,'op-'+e.id+'.json');if(seen.has(e.id))throw error('操作番号が重複しています。','integrity');seen.add(e.id);}
  const prefix=previous&&previous.ordered.length<=ordered.length&&previous.ordered.every((e,i)=>ordered[i].id===e.id);
  const rejected=prefix?previous.rejected.slice():[],results=prefix?new Map(previous.results):new Map();let db=prefix?previous.db:DeskCore.empty();
  for(let i=prefix?previous.ordered.length:0;i<ordered.length;i++){const e=ordered[i];try{const r=apply(db,e);db=r.db;results.set(e.id,r.result);}catch(err){if(err.code!=='conflict')throw err;rejected.push({id:e.id,kind:e.action.kind,ids:err.ids||[],message:err.message,stamp:e.stamp});}}
  return{db,rejected,results,ordered,maximum:ordered.at(-1)?.clock||0,signature:ordered.map(e=>e.id).join(',')};
 }
 let dbPromise;
 const prefs={
  open(){if(!dbPromise)dbPromise=new Promise((resolve,reject)=>{const r=indexedDB.open('mention-desk-folder-v1',1);r.onupgradeneeded=()=>r.result.createObjectStore('settings');r.onerror=()=>reject(r.error);r.onblocked=()=>reject(error('他の受付ボードを閉じて、設定を再度開いてください。'));r.onsuccess=()=>resolve(r.result);});return dbPromise;},
  async get(key){const db=await this.open();return new Promise((resolve,reject)=>{const tx=db.transaction('settings','readonly'),r=tx.objectStore('settings').get(key);r.onsuccess=()=>resolve(r.result);r.onerror=()=>reject(r.error);});},
  async set(key,value){const db=await this.open();return new Promise((resolve,reject)=>{const tx=db.transaction('settings','readwrite');tx.objectStore('settings').put(value,key);tx.oncomplete=()=>resolve();tx.onabort=tx.onerror=()=>reject(tx.error||error('設定を保存できません。'));});}
 };
 async function optional(folder,kind,name){try{return await folder[kind==='file'?'getFileHandle':'getDirectoryHandle'](name);}catch(e){if(e.name==='NotFoundError')return null;throw e;}}
 class Client{
  constructor({storage=localStorage,settings=prefs,onUpdate=()=>{},onStatus=()=>{},settleMs=250,today=dayKey}={}){Object.assign(this,{storage,settings,onUpdate,onStatus,settleMs,today,day:today(),contexts:new Map(),root:null,key:null,cache:new Map(),state:replay([]),queue:Promise.resolve(),poll:null,integrityError:null,checkCursor:0,pendingCount:0});if(!validDay(this.day))throw error('PCの日付を確認してください。','invalid');}
  useDay(day){if(day!==null&&!validDay(day))throw error('保存先の日付が不正です。','integrity');if(this.day===day)return;this.contexts.set(this.day,{cache:this.cache,state:this.state,integrityError:this.integrityError,checkCursor:this.checkCursor});Object.assign(this,this.contexts.get(day)||{cache:new Map(),state:replay([]),integrityError:null,checkCursor:0},{day});}
  async directory(create=false){const base=create?await this.root.getDirectoryHandle(DIR,{create:true}):await optional(this.root,'directory',DIR);if(!base||this.day===null)return base;return create?await base.getDirectoryHandle(this.day,{create:true}):await optional(base,'directory',this.day);}
  dayError(){return this.day===this.today()?null:error('日付が変わりました。「共有を再接続」して当日分のCSVを集計してください。','daychange');}
  checkDay(){const problem=this.dayError();if(problem)throw problem;}
  serial(fn){const run=()=>typeof navigator!=='undefined'&&navigator.locks?navigator.locks.request('mention-desk-folder-queue',fn):fn();const next=this.queue.then(run,run);this.queue=next.catch(()=>{});return next;}
  jobs(){const list=[];for(let i=0;i<this.storage.length;i++){const k=this.storage.key(i);const legacy=k?.startsWith(LEGACY_PREFIX);if(!legacy&&!k?.startsWith(PREFIX))continue;let j;try{j=JSON.parse(this.storage.getItem(k));}catch{throw error('端末内の保存待ちデータが壊れています。消去せず確認してください。','integrity');}if(!j||!ID.test(j.id)||k!==(legacy?LEGACY_PREFIX:PREFIX)+j.id||typeof j.rootKey!=='string'||(legacy?j.day!==undefined:!validDay(j.day))||!j.action||!Number.isFinite(Date.parse(j.stamp)))throw error('端末内の保存待ちデータの形式が不正です。','integrity');validateAction(j.action);if(j.event)validateEvent(j.event,'op-'+j.id+'.json');Object.defineProperty(j,'legacy',{value:!!legacy,enumerable:false});list.push(j);}return list.sort((a,b)=>cmp(a.stamp,b.stamp)||cmp(a.id,b.id));}
  keep(job){try{this.storage.setItem((job.legacy?LEGACY_PREFIX:PREFIX)+job.id,JSON.stringify(job));}catch{throw error('操作を端末内に退避できません。共有への書き込みは開始していません。ブラウザーの空き容量を確認してください。');}}
  remove(job){this.storage.removeItem((job.legacy?LEGACY_PREFIX:PREFIX)+job.id);}
  async belongs(job,root=this.root,key=this.key){if(job.rootKey===key&&this.root&&await root.isSameEntry(this.root))return true;const saved=await this.settings.get(job.rootKey);return!!saved&&await root.isSameEntry(saved);}
  async pending(){let count=0;for(const j of this.jobs())if(await this.belongs(j))count++;return count;}
  async permission(){if(!this.root||await this.root.queryPermission({mode:'readwrite'})!=='granted')throw error('共有フォルダの読み書き許可が必要です。「共有を接続」で再許可してください。','permission');}
  async connect(handle,key){return this.serial(async()=>{
   if(await handle.queryPermission({mode:'readwrite'})!=='granted')throw error('共有フォルダの読み書きが許可されていません。','permission');
   for(const j of this.jobs())if(!await this.belongs(j,handle,key))throw error('別の共有先への保存待ちがあります。元の共有先へ接続して保存してから変更してください。');
   if(!key)key='root-'+uid();await this.settings.set(key,handle);
   const day=this.today();if(!validDay(day))throw error('PCの日付を確認してください。','invalid');
   const old={root:this.root,key:this.key,day:this.day,contexts:this.contexts,cache:this.cache,state:this.state,integrityError:this.integrityError,checkCursor:this.checkCursor,pendingCount:this.pendingCount};
   this.root=handle;this.key=key;this.day=day;this.contexts=new Map();this.cache=new Map();this.state=replay([]);this.integrityError=null;this.checkCursor=0;
   try{
    // Test actual write/close/readback in the pinned date folder.
    const directory=await this.directory(true),name='probe-'+uid()+'.tmp',probe=await directory.getFileHandle(name,{create:true});let stream;
    try{stream=await probe.createWritable();await stream.write('MentionDesk write check');await stream.close();if(await(await probe.getFile()).text()!=='MentionDesk write check')throw error('共有先への書き込み結果を確認できません。');}catch(e){try{await stream?.abort();}catch{}throw e;}finally{try{await directory.removeEntry(name);}catch{}}
    await this.scan();await this.settings.set('current',{handle,key});let recovered=null;try{await this.drain();}catch(e){if(e.code==='conflict')recovered=e;else throw e;}this.checkDay();this.publish(recovered);return this.state;
   }catch(e){Object.assign(this,old);throw e;}
  });}
  async scan(){
   if(this.integrityError)throw this.integrityError;await this.permission();
   const directory=await this.directory(),folder=directory?await optional(directory,'directory','operations'):null,seen=new Set(),names=[...this.cache.keys()],checks=new Set();
   for(let i=0;i<Math.min(16,names.length);i++)checks.add(names[(this.checkCursor+i)%names.length]);this.checkCursor=names.length?(this.checkCursor+checks.size)%names.length:0;
   try{
    if(folder)for await(const [name,handle]of folder.entries()){
     if(handle.kind!=='file'||!FILE.test(name))continue;seen.add(name);if(seen.size>50000)throw error('共有操作が50,000件を超えています。履歴を消さずに管理担当へ確認してください。','integrity');
     const cached=this.cache.get(name);if(cached&&!checks.has(name))continue;
     const file=await handle.getFile();if(!file.size){if(cached)throw error('保存済みの共有履歴が空になっています：'+name,'integrity');continue;}
     if(file.size>2*1024*1024)throw error('共有履歴のファイルが大きすぎます：'+name,'integrity');
     const text=await file.text(),mark=file.size+':'+file.lastModified;
     if(cached){if(cached.mark!==mark||cached.text!==text)throw error('保存済みの共有履歴が変更されています：'+name,'integrity');continue;}
     let event;try{event=JSON.parse(text);}catch{throw error('共有履歴が壊れています：'+name,'integrity');}validateEvent(event,name);this.cache.set(name,{event,mark,text});
    }
    for(const name of this.cache.keys())if(!seen.has(name))throw error('保存済みの共有履歴が削除されています：'+name,'integrity');
    // Fully replay when a late file changes order; never rely on PC wall clocks.
    const signature=[...this.cache.values()].map(x=>x.event).sort((a,b)=>a.clock-b.clock||cmp(a.id,b.id)).map(x=>x.id).join(',');
    if(this.state.signature!==signature)this.state=replay([...this.cache.values()].map(x=>x.event),this.state);return this.state;
   }catch(e){if(e.code==='integrity')this.integrityError=e;throw e;}
  }
  publish(problem=null){problem=problem||this.dayError();this.onUpdate({...this.state,day:this.day});this.onStatus({connected:!!this.root,folder:this.root?.name||'',day:this.day,pending:this.pendingCount,rejected:this.state.rejected,error:problem});}
  async write(job){
   if((job.day??null)!==this.day)throw error('保存待ちの操作と日付フォルダが一致しません。','integrity');
   const directory=await this.directory(true),folder=await directory.getDirectoryHandle('operations',{create:true}),name='op-'+job.id+'.json',text=JSON.stringify(job.event,null,2)+'\n';
   let handle=await optional(folder,'file',name);if(handle){const file=await handle.getFile();if(file.size){if(await file.text()!==text)throw error('同じ操作番号に別の内容があります。上書きを停止しました。','integrity');return;}}
   else handle=await folder.getFileHandle(name,{create:true});const stream=await handle.createWritable();
   try{await stream.write(text);await stream.close();}catch(e){try{await stream.abort();}catch{}throw e;}
   if(await(await handle.getFile()).text()!==text)throw error('共有への保存結果を確認できません。同じ操作番号で再試行します。');
  }
  async finish(job){
   if((job.day??null)!==this.day)throw error('保存待ちの操作と日付フォルダが一致しません。','integrity');
   await this.scan();
   // An already saved event is replayed and checked before any retry writes.
   let rejected=this.state.rejected.find(x=>x.id===job.id);
   if(rejected){this.remove(job);throw error(rejected.message,'conflict',rejected.ids);}
   if(!this.state.results.has(job.id)){
    if(!job.event){
     // Preflight includes the exact claim/note tokens from the displayed snapshot.
     const e={schemaVersion:1,application:APP,id:job.id,clock:this.state.maximum+1,stamp:job.stamp,action:job.action};
     try{apply(this.state.db,e);}catch(err){if(err.code==='conflict')this.remove(job);throw err;}
     job.event=e;this.keep(job);
    }
    // If an earlier attempt did not leave a committed file, check again before
    // retrying its prepared event. Do not revive an offline stale claim.
    const directory=await this.directory(),folder=directory?await optional(directory,'directory','operations'):null;
    const prior=folder?await optional(folder,'file','op-'+job.id+'.json'):null,committed=prior?(await prior.getFile()).size>0:false;
    if(!committed){try{apply(this.state.db,job.event);}catch(err){if(err.code==='conflict')this.remove(job);throw err;}}
    await this.write(job);await this.scan();
    // Short additional read narrows the usual concurrent-click window.
    // Visibility delays can exceed this wait, so UI explicitly reports late conflicts.
    if(this.settleMs)await new Promise(resolve=>setTimeout(resolve,this.settleMs));await this.scan();
   }
   rejected=this.state.rejected.find(x=>x.id===job.id);if(rejected){this.remove(job);throw error(rejected.message,'conflict',rejected.ids);}
   if(!this.state.results.has(job.id))throw error('共有先への反映を確認できません。保存待ちを保持しています。');
   const result=this.state.results.get(job.id);this.remove(job);return result;
  }
  async drain(){
   const day=this.day;let problem=null;
   try{for(const j of this.jobs())if(await this.belongs(j)){this.useDay(j.day??null);try{await this.finish(j);}catch(e){problem=e;if(e.code!=='conflict')break;}}}
   finally{if(this.day!==day){this.useDay(day);try{await this.scan();}catch(e){problem=problem||e;}}this.pendingCount=await this.pending();}
   if(problem)throw problem;
  }
  async submit(action){return this.serial(async()=>{
   try{
    validateAction(action);if(this.day===null||action.kind!=='note')this.checkDay();await this.permission();
    let retry=null;for(const j of this.jobs())if((j.day??null)===this.day&&await this.belongs(j)&&JSON.stringify(j.action)===JSON.stringify(action)){retry=j;break;}
    if(retry){const result=await this.finish(retry);this.pendingCount=await this.pending();this.publish();return{result,db:this.state.db,eventId:retry.id};}
    await this.drain();
    const job={id:uid(),rootKey:this.key,day:this.day,stamp:new Date().toISOString(),action:clone(action)};this.keep(job);this.pendingCount=await this.pending();this.publish();
    const result=await this.finish(job);this.pendingCount=await this.pending();this.publish();return{result,db:this.state.db,eventId:job.id};
   }catch(e){this.pendingCount=await this.pending().catch(()=>1);this.publish(e);throw e;}
  });}
  async refresh(){return this.serial(async()=>{try{await this.permission();await this.drain();await this.scan();this.checkDay();this.publish();return this.state;}catch(e){this.publish(e);throw e;}});}
 }
 function chunks(records){const result=[];let batch=[],size=0;for(const r of records){const length=new TextEncoder().encode(JSON.stringify(r)).length;if(length>900000)throw error('1件のCSVデータが大きすぎます。内容を確認してください。','invalid');if(batch.length&&(batch.length>=500||size+length>900000)){result.push(batch);batch=[];size=0;}batch.push(r);size+=length;}if(batch.length)result.push(batch);return result;}
 return{Client,prefs,replay,apply,validateEvent,validateAction,chunks,uid,DIR,dayKey};
})();

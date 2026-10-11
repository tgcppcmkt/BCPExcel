import test from 'node:test';
import assert from 'node:assert/strict';
import fs from 'node:fs/promises';
import vm from 'node:vm';
import os from 'node:os';
import path from 'node:path';

const html = await fs.readFile(new URL('../TaskBoard_v1.0.html',import.meta.url),'utf8');
const scripts = [...html.matchAll(/<script[^>]*>([\s\S]*?)<\/script>/g)].map(m=>m[1]);
for(const script of scripts) new vm.Script(script);
const sandbox={};
vm.createContext(sandbox);
vm.runInContext(scripts[0],sandbox);
const C=sandbox.TaskBoardCore;
const normalized=value=>JSON.parse(JSON.stringify(value));
const op=(id,taskId,seq,patch)=>({v:1,id,taskId,seq,patch});
const task=(time='10:00',extra={})=>({date:'2026-10-11',time,title:'進捗確認',assignee:'田中',notes:'確認事項',status:'pending',...extra});

test('時間・日付・タスク形式を検証し、範囲外のデータを拒否する',()=>{
  assert.equal(C.validTime('09:05'),true);
  for(const time of ['9:05','24:00','12:60','not-a-time'])assert.equal(C.validTime(time),false);
  assert.equal(C.validDate('2028-02-29'),true);
  for(const date of ['2026-02-29','2026-04-31','2026-13-01','2026-00-01'])assert.equal(C.validDate(date),false);
  assert.equal(C.validOperation(op('event0001','task00001',1,task())),true);
  assert.equal(C.validOperation(op('event0001','task00001',1,{unexpected:'field'})),false);
  assert.equal(C.validOperation(op('event0001','task00001',1,{title:'   '})),false);
});

test('各列の並び順は予定時刻で決まり、同時刻でも安定する',()=>{
  for(const status of C.STATUSES) {
    const tasks=['13:00','09:30','11:00'].map((time,i)=>({...task(time,{status}),id:'task0000'+i}));
    assert.deepEqual(normalized(C.sortedTasks(tasks)).map(t=>t.time),['09:30','11:00','13:00']);
  }
  assert.deepEqual(normalized(C.sortedTasks([{...task(),id:'task00002'},{...task(),id:'task00001'}])).map(t=>t.id),['task00001','task00002']);
});

test('予定時刻を過ぎた未対応だけを通知し、前日・未来日も正しく判定する',()=>{
  const exact=new Date('2026-10-11T10:00:00').getTime();
  assert.equal(C.isOverdue(task(),exact),false);
  assert.equal(C.isOverdue(task(),exact+1),true);
  assert.equal(C.isOverdue(task('09:00',{status:'progress'}),exact),false);
  assert.equal(C.isOverdue(task('09:00',{status:'done'}),exact),false);
  assert.equal(C.isOverdue(task('23:00',{date:'2026-10-10'}),exact),true);
  assert.equal(C.isOverdue(task('09:00',{date:'2026-10-12'}),exact),false);
});

test('別々の項目の同時編集を保持し、受信順が違っても同じ結果になる',()=>{
  const events=[
    op('event0001','task00001',1,task()),
    op('event0002','task00001',2,{status:'progress'}),
    op('event0003','task00001',2,{notes:'別の担当者がメモを更新'}),
    op('event0004','task00001',3,{status:'done'})
  ];
  const a=normalized(C.reduceOperations(events)),b=normalized(C.reduceOperations([...events].reverse()));
  assert.deepEqual(a,b);
  assert.equal(a[0].status,'done');
  assert.equal(a[0].notes,'別の担当者がメモを更新');
  assert.equal(C.nextSequence(events),4);
});

test('同じ項目の同時変更を一意に解決し、削除後の遅延編集で復活しない',()=>{
  const events=[
    op('event0001','task00001',1,task()),
    op('event0002','task00001',2,{status:'progress'}),
    op('event0003','task00001',2,{status:'done'})
  ];
  assert.equal(C.reduceOperations(events)[0].status,'done');
  assert.deepEqual(normalized(C.reduceOperations(events)),normalized(C.reduceOperations([...events].reverse())));
  events.push(op('event0004','task00001',3,{deleted:true}),op('event0005','task00001',4,{notes:'削除前に開いたフォームから保存'}));
  assert.equal(C.reduceOperations(events).length,0);
});

// Node file handles implement the browser handle contract against real temp files.
// Writes become visible only on close, as with FileSystemWritableFileStream.
class DirectoryHandle {
  constructor(directory){this.directory=directory;this.kind='directory';}
  async *entries(){for(const name of await fs.readdir(this.directory))yield[name,new FileHandle(path.join(this.directory,name))];}
  async getFileHandle(name,{create}={}) {
    const target=path.join(this.directory,name);
    if(create)await fs.writeFile(target,'',{flag:'a'});
    return new FileHandle(target);
  }
}
class FileHandle {
  constructor(filename){this.filename=filename;this.kind='file';}
  async getFile(){const data=await fs.readFile(this.filename,'utf8');return{size:Buffer.byteLength(data),text:async()=>data};}
  async createWritable(){
    let buffer='',aborted=false;
    return{
      write:async data=>{buffer=data;},
      close:async()=>{assert.equal(aborted,false);const tmp=this.filename+'.tmp-'+Math.random().toString(36).slice(2);await fs.writeFile(tmp,buffer);await fs.rename(tmp,this.filename);},
      abort:async()=>{aborted=true;}
    };
  }
}

test('共有フォルダで2人が同時に追加・編集しても、別タスクや別項目が消えない',async()=>{
  const directory=await fs.mkdtemp(path.join(os.tmpdir(),'taskboard-'));
  try {
    const a=new C.FolderStore(new DirectoryHandle(directory)),b=new C.FolderStore(new DirectoryHandle(directory));
    await Promise.all([a.refresh(),b.refresh()]);
    await Promise.all([
      a.append(op('event0001','task00001',1,task('09:00'))),
      b.append(op('event0002','task00002',1,task('11:00')))
    ]);
    await Promise.all([a.refresh(),b.refresh()]);
    assert.equal(a.operations.size,2);assert.equal(b.operations.size,2);
    assert.equal(C.reduceOperations(a.operations.values()).length,2);
    await Promise.all([
      a.append(op('event0003','task00001',2,{status:'progress'})),
      b.append(op('event0004','task00001',2,{notes:'サポート投入を確認'}))
    ]);
    await Promise.all([a.refresh(),b.refresh()]);
    assert.deepEqual(normalized(C.sortedTasks(C.reduceOperations(a.operations.values()))),normalized(C.sortedTasks(C.reduceOperations(b.operations.values()))));
    const first=C.reduceOperations(a.operations.values()).find(t=>t.id==='task00001');
    assert.equal(first.status,'progress');assert.equal(first.notes,'サポート投入を確認');
    assert.equal((await fs.readdir(directory)).filter(n=>n.endsWith('.json')).length,4);
  }finally{await fs.rm(directory,{recursive:true,force:true});}
});

test('書き込み途中の共有ファイルを次回の同期で再読込し、欠落扱いにしない',async()=>{
  const directory=await fs.mkdtemp(path.join(os.tmpdir(),'taskboard-'));
  try {
    const store=new C.FolderStore(new DirectoryHandle(directory));
    const target=path.join(directory,'op-event0001.json');
    await fs.writeFile(target,'{');
    assert.equal(await store.refresh(),0);assert.equal(store.unread,1);
    await fs.writeFile(target,JSON.stringify(op('event0001','task00001',1,task())));
    assert.equal(await store.refresh(),1);assert.equal(store.unread,0);assert.equal(store.operations.size,1);
    assert.equal(await store.refresh(),0);
  }finally{await fs.rm(directory,{recursive:true,force:true});}
});

test('保存失敗時は、共有済みとして処理しない',async()=>{
  const handle={
    entries:async function*(){},
    getFileHandle:async()=>({createWritable:async()=>({write:async()=>{throw new Error('network offline');},abort:async()=>{}})})
  };
  const store=new C.FolderStore(handle);
  await assert.rejects(store.append(op('event0001','task00001',1,task())),/network offline/);
  assert.equal(store.operations.size,0);
});

test('画面のID参照がすべて実在し、入力欄・ダイアログ・カラムが揃う',()=>{
  const ids=[...html.matchAll(/\bid="([^"]+)"/g)].map(m=>m[1]);
  assert.equal(new Set(ids).size,ids.length);
  const references=[...scripts[1].matchAll(/\$\('([^']+)'\)/g)].map(m=>m[1]);
  for(const id of references)assert.ok(ids.includes(id),'Missing UI element '+id);
  for(const status of C.STATUSES)assert.ok(ids.includes('list-'+status));
  assert.match(html,/@media\(prefers-color-scheme:dark\)/);
  assert.match(html,/@media\(prefers-reduced-motion:reduce\)/);
  assert.match(html,/<html lang="ja">/);
  assert.equal((html.match(/<dialog\b/g)||[]).length,4);
});


test('自分が担当の検索は担当欄で一致させ、タスク名や備考の名前で誤検出しない',()=>{
  const tasks=[
    {...task('09:00',{assignee:'田口'}),id:'task00001'},
    {...task('10:00',{assignee:'田口太郎'}),id:'task00002'},
    {...task('11:00',{assignee:'佐藤',title:'田口へ連絡'}),id:'task00003'},
    {...task('12:00',{assignee:'鈴木',notes:'田口が確認予定'}),id:'task00004'},
    {...task('13:00',{assignee:''}),id:'task00005'}
  ];
  assert.deepEqual(normalized(C.filterTasks(tasks,'','田口',true)).map(t=>t.id),['task00001']);
  assert.deepEqual(normalized(C.filterTasks(tasks,'田口','田口',false)).map(t=>t.id),['task00001','task00002','task00003','task00004']);
  assert.equal(C.filterTasks(tasks,'','',true).length,0);
  assert.equal(C.filterTasks(tasks,'','田口',false).length,5);
});

test('自分が担当とキーワード検索を組み合わせ、解除すれば他の担当も検索できる',()=>{
  const tasks=[
    {...task('09:00',{assignee:'田口',title:'日程調整の進捗確認'}),id:'task00001'},
    {...task('10:00',{assignee:'田口',title:'引き継ぎメモ更新'}),id:'task00002'},
    {...task('11:00',{assignee:'佐藤',title:'日程調整の進捗確認'}),id:'task00003'}
  ];
  const original=JSON.stringify(tasks);
  assert.deepEqual(normalized(C.filterTasks(tasks,'日程','田口',true)).map(t=>t.id),['task00001']);
  assert.deepEqual(normalized(C.filterTasks(tasks,'日程','田口',false)).map(t=>t.id),['task00001','task00003']);
  assert.equal(JSON.stringify(tasks),original);
});

test('担当名の空白・全半角を吸収し、登録名の変更も絞り込みに反映する',()=>{
  const tasks=[
    {...task('09:00',{assignee:'田口　太郎'}),id:'task00001'},
    {...task('10:00',{assignee:'Ｓａｔｏ'}),id:'task00002'}
  ];
  assert.equal(C.isAssignedToMe(tasks[0],' 田口 太郎 '),true);
  assert.equal(C.isAssignedToMe(tasks[1],'sato'),true);
  assert.deepEqual(normalized(C.filterTasks(tasks,'','田口太郎',true)).map(t=>t.id),['task00001']);
  assert.deepEqual(normalized(C.filterTasks(tasks,'','Sato',true)).map(t=>t.id),['task00002']);
});

test('初回登録する名前を検証し、空欄・空白だけ・文字数超過を拒否する',()=>{
  assert.equal(C.validPersonalName('田口'),true);
  assert.equal(C.validPersonalName(' 田口 '),true);
  assert.equal(C.validPersonalName('名'.repeat(100)),true);
  for(const name of ['', '　 ', null, '名'.repeat(101)])assert.equal(C.validPersonalName(name),false);
});

test('1名の過去データを保持し、任意の担当者②を検証して保存する',()=>{
  const legacy=op('event0001','task00001',1,task());
  assert.equal(C.validOperation(legacy),true);
  assert.equal(C.reduceOperations([legacy])[0].assignee,'田中');
  assert.equal(C.reduceOperations([legacy])[0].assignee2,'');
  for(const name of ['', '佐藤', '名'.repeat(100)]) {
    const event=op('event0002','task00002',1,task('11:00',{assignee2:name}));
    assert.equal(C.validOperation(event),true);
    const restored=C.reduceOperations(JSON.parse(JSON.stringify([event])))[0];
    assert.equal(restored.assignee,'田中');
    assert.equal(restored.assignee2,name);
  }
  for(const name of [null,undefined,7,'名'.repeat(101)])
    assert.equal(C.validOperation(op('event0002','task00001',2,{assignee2:name})),false);
});

test('担当者②も自分のタスク・キーワード検索の対象になり、重複表示しない',()=>{
  const tasks=[
    {...task('09:00',{assignee:'田口',assignee2:'佐藤',title:'日程調整'}),id:'task00001'},
    {...task('10:00',{assignee:'佐藤',assignee2:'田口',title:'日程確認'}),id:'task00002'},
    {...task('11:00',{assignee:'田口',assignee2:'田口',title:'引き継ぎ'}),id:'task00003'},
    {...task('12:00',{assignee:'',assignee2:'田口'}),id:'task00004'},
    {...task('13:00',{assignee:'佐藤',assignee2:'田口太郎'}),id:'task00005'},
    {...task('14:00',{assignee:'鈴木',notes:'田口が確認'}),id:'task00006'}
  ];
  assert.deepEqual(normalized(C.filterTasks(tasks,'','田口',true)).map(t=>t.id),['task00001','task00002','task00003','task00004']);
  assert.deepEqual(normalized(C.filterTasks(tasks,'日程','田口',true)).map(t=>t.id),['task00001','task00002']);
  assert.deepEqual(normalized(C.filterTasks(tasks,'田口太郎','田口',false)).map(t=>t.id),['task00005']);
  assert.equal(C.filterTasks(tasks,'田口太郎','田口',true).length,0);
  assert.equal(C.isAssignedToMe(task('10:00',{assignee:'田中',assignee2:' Ｓａｔｏ　太郎 '}),'sato 太郎'),true);
});

test('2名の担当欄を独立して同時編集でき、担当者②を消しても①は残る',()=>{
  const events=[
    op('event0001','task00001',1,task()),
    op('event0002','task00001',2,{assignee:'鈴木'}),
    op('event0003','task00001',2,{assignee2:'佐藤'}),
    op('event0004','task00001',3,{status:'progress'})
  ];
  const reduced=normalized(C.reduceOperations(events));
  assert.deepEqual(reduced,normalized(C.reduceOperations([...events].reverse())));
  assert.equal(reduced[0].assignee,'鈴木');
  assert.equal(reduced[0].assignee2,'佐藤');
  assert.equal(reduced[0].status,'progress');
  events.push(op('event0005','task00001',4,{assignee2:''}));
  assert.equal(C.reduceOperations(events)[0].assignee,'鈴木');
  assert.equal(C.reduceOperations(events)[0].assignee2,'');
});

test('バックアップ復元で2名の担当を保持し、過去の1名バックアップも取り込める',()=>{
  const original=task('10:00',{assignee2:'佐藤',id:'old-task-id',demo:true});
  const patch=normalized(C.backupTaskPatch(JSON.parse(JSON.stringify(original))));
  assert.equal(patch.assignee,'田中');
  assert.equal(patch.assignee2,'佐藤');
  assert.equal(patch.id,undefined);
  assert.equal(patch.demo,undefined);
  const restored=C.reduceOperations([op('event0001','task00001',1,patch)])[0];
  assert.equal(restored.assignee2,'佐藤');
  assert.equal(C.backupTaskPatch(task()).assignee2,'');
  for(const invalid of [null,[],{},task('10:00',{assignee2:33}),task('10:00',{assignee2:null})])
    assert.throws(()=>C.backupTaskPatch(invalid),/不正/);
});

test('共有フォルダの再読込で2名の担当者を保持する',async()=>{
  const directory=await fs.mkdtemp(path.join(os.tmpdir(),'taskboard-'));
  try {
    const writer=new C.FolderStore(new DirectoryHandle(directory));
    await writer.append(op('event0001','task00001',1,task('09:00',{assignee2:'佐藤'})));
    const reader=new C.FolderStore(new DirectoryHandle(directory));
    assert.equal(await reader.refresh(),1);
    const restored=C.reduceOperations(reader.operations.values())[0];
    assert.equal(restored.assignee,'田中');
    assert.equal(restored.assignee2,'佐藤');
    assert.equal(C.isAssignedToMe(restored,'佐藤'),true);
  }finally{await fs.rm(directory,{recursive:true,force:true});}
});


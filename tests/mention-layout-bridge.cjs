// Exercise the actual async bridge with local file/process boundaries replaced.
const fs=require('node:fs'),vm=require('node:vm'),path=require('node:path'),assert=require('node:assert/strict');
const source=fs.readFileSync(path.join(__dirname,'../メンション依頼フォーム_HTA_最新版/mention-layout.js'),'utf8');
const files=new Map(),commands=[],intervals=new Map();
let failures=0,clock=0,next=0,view={width:600,height:400},outer={width:620,height:440};
const fso={
    BuildPath:(a,b)=>a+'\\'+b, GetSpecialFolder:()=> 'C:\\Temp', CreateFolder:()=>{},
    FileExists:p=>files.has(p)||/mention-layout-host\.(ps1|cs)$|powershell\.exe$/.test(p),
    CreateTextFile:p=>({Write:v=>files.set(p,v),Close(){}}),
    OpenTextFile:p=>({ReadAll:()=>files.get(p),Close(){}})
};
const c=vm.createContext({
    Date:class {getTime(){return clock;}}, Math, document:{title:'メンション依頼フォーム'},
    window:{setInterval:f=>{intervals.set(++next,f);return next;},clearInterval:id=>intervals.delete(id)},
    ActiveXObject:function(name){return name==='Scripting.FileSystemObject'?fso:{ExpandEnvironmentStrings:()=> 'C:\\Windows',Run:cmd=>commands.push(cmd)};},
    getCurrentFolderPath:()=> 'C:\\日本語 フォルダ', getViewportSize:()=>view,getWindowOuterSize:()=>outer,
    layoutAdjustmentUnavailable:()=>failures++
});
vm.runInContext(source,c);
function tick(){clock+=100; for(const fn of intervals.values()) fn();}
function getFile(name){return [...files.keys()].find(p=>p.endsWith('\\'+name));}
assert.equal(c.MentionLayout.request(1548,400,true),true);
assert.equal(commands.length,1);
assert.ok(commands[0].includes('-File "C:\\日本語 フォルダ\\mention-layout-host.ps1"'));
assert.ok(!commands[0].includes('ExecutionPolicy'));
const folder=getFile('heartbeat.txt').replace(/\\heartbeat.txt$/,'');
files.set(folder+'\\ready.txt','OK'); tick();
let r=files.get(folder+'\\request.txt').split('|');
assert.equal(r[2],'400'); assert.equal(r[3],'600');
c.MentionLayout.request(1548,700,false);
c.MentionLayout.request(1548,1000,false);
tick(); assert.equal(files.get(folder+'\\request.txt').split('|')[0],'1');
view={width:1500,height:360};outer={width:1520,height:400};
files.set(folder+'\\result.txt','1|OK');tick();
r=files.get(folder+'\\request.txt').split('|');
assert.equal(r[0],'2');assert.equal(r[2],'1000');assert.equal(r[3],'1500');
assert.equal(commands.length,1,'one helper per form');
files.set(folder+'\\result.txt','2|SKIP');tick();
assert.equal(failures,0);
files.set(folder+'\\error.txt','Native call failed');tick();
assert.equal(failures,1);assert.equal(files.get(folder+'\\stop.txt'),'STOP');
assert.equal(c.MentionLayout.request(1548,700,false),false);
files.clear();intervals.clear();commands.length=0;failures=0;
vm.runInContext(source,c);
assert.equal(c.MentionLayout.request(1548,400,true),true);
for(let i=0;i<52;i++)tick();
assert.equal(failures,1,'a blocked PowerShell launch must fall back after timeout');
assert.equal(c.MentionLayout.request(1548,700,false),false);
assert.equal(commands.length,1,'do not repeatedly relaunch a blocked helper');
console.log('Bridge passed: launch path, one worker, serialized/coalesced requests, fresh viewport, minimized result, error and blocked-start fallback.');

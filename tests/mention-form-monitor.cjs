// node tests/mention-form-monitor.cjs [optional source path]
// Runs the real layout/card handlers with simulated HTA window/DOM boundaries.
// Does not access ActiveX, write CSVs, or send requests. Windows/DPI QA is separate.
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const assert = require('node:assert/strict');
const source = fs.readFileSync(process.argv[2] || path.join(__dirname, '..',
    'メンション依頼フォーム_HTA_最新版', 'mention-form.js'), 'utf8');

function fixture(screenProperties = {}) {
    let clock = 0, sequence = 0;
    const timers = new Map(), intervals = new Map(), elements = {};
    const moves = [], sizes = [];
    const screen = {availWidth:1920, availHeight:1040, ...screenProperties};
    const win = {
        screenLeft:100, screenTop:120, outerWidth:1762, outerHeight:370,
        setTimeout(fn, delay) {const id = ++sequence; timers.set(id, {fn, at:clock+delay}); return id;},
        clearTimeout(id) {timers.delete(id);},
        setInterval(fn) {const id = ++sequence; intervals.set(id, fn); return id;},
        clearInterval(id) {intervals.delete(id);},
        moveTo(x,y) {moves.push([x,y]); win.screenLeft=x; win.screenTop=y;},
        moveBy(x,y) {win.moveTo(win.screenLeft+x, win.screenTop+y);},
        resizeTo(w,h) {
            sizes.push([w,h]); win.outerWidth=w; win.outerHeight=h;
            if (win.resizeShift) {win.screenLeft += win.resizeShift[0]; win.screenTop += win.resizeShift[1];}
        },
        scrollTo() {}
    };
    function el(id) {
        if (!elements[id]) elements[id] = {id, style:{}, value:'', checked:false,
            className: /request[23]/.test(id) ? 'request-card hidden' : '',
            focus() {}, setAttribute() {}, removeAttribute() {}};
        return elements[id];
    }
    const document = {getElementById:el, getElementsByTagName:()=>[], body:{style:{}}, documentElement:{}};
    Object.defineProperties(document.documentElement, {
        clientWidth:{get:()=>win.outerWidth-54}, clientHeight:{get:()=>win.outerHeight-76}
    });
    const c = vm.createContext({window:win, screen, document});
    vm.runInContext(source, c);
    Object.defineProperties(el('app'), {
        offsetWidth:{get:()=>1680}, offsetHeight:{get:()=>220*c.visibleRequestCount+40}
    });
    // Valid request input, leaving actual show/cancel/reset and resize handlers in place.
    for (let n=1; n<=3; n++) {
        for (const field of ['org','ca','mailMemo','processedAt','type']) el(field+n).value='TEST';
        el('attendanceWork'+n).checked=true;
    }
    function advance(ms=100) {
        const end=clock+ms;
        for (let steps=0; steps<100; steps++) {
            const next=[...timers].sort((a,b)=>a[1].at-b[1].at)[0];
            if (!next || next[1].at>end) {clock=end; return;}
            timers.delete(next[0]); clock=next[1].at; next[1].fn();
        }
        throw new Error('layout timer loop');
    }
    function place(x,y) {win.screenLeft=x; win.screenTop=y; moves.length=0;}
    const position = () => [win.screenLeft,win.screenTop];
    c.resizeApp(true); advance();
    return {c,win,screen,el,moves,sizes,advance,place,position,
        tickWatch:()=>{for (const fn of intervals.values()) fn(); advance();}};
}

const cases=[];
function test(name, fn) {cases.push([name,fn]);}

for (const [label, props] of [
    ['missing origin', {}],
    ['null origin', {availLeft:null,availTop:null}],
    ['empty origin', {availLeft:'',availTop:''}],
    ['stale primary origin', {availLeft:0,availTop:0}]
]) {
    for (const [side, x,y] of [['right',2040,150],['left',-1880,150],['above',100,-1050],['below',100,1220]]) {
        test(`${label}, external ${side}: expand/cancel/reset preserve position`, () => {
            const f=fixture(props); f.place(x,y);
            f.c.showRequest(2); f.advance();
            assert.equal(f.c.visibleRequestCount,2);
            assert.deepEqual(f.position(),[x,y], 'request 2 moved to another monitor');
            const twoHeight=f.win.outerHeight;
            f.c.showRequest(3); f.advance();
            assert.equal(f.c.visibleRequestCount,3);
            assert.equal(f.win.outerHeight,twoHeight, 'without native geometry, keep the visible viewport and use scrollbars');
            assert.deepEqual(f.position(),[x,y], 'request 3 moved');
            f.c.cancelRequest(3); f.advance();
            f.c.cancelRequest(2); f.advance();
            f.c.resetAfterSend(); f.advance();
            assert.deepEqual(f.position(),[x,y]);
            assert.equal(f.el('app').style.visibility,'visible');
            assert.equal(f.moves.length,0, 'unknown/stale origins must not trigger relocation');
        });
    }
}

test('valid external work area: keep upper-center fitting within that monitor', () => {
    const f=fixture(); Object.assign(f.screen,{availLeft:1920,availTop:0});
    f.place(2050,850); f.c.showRequest(2); f.advance();
    assert.ok(f.position()[0]>=1920 && f.position()[0]+f.win.outerWidth<=3840);
    assert.ok(f.position()[1]<850 && f.position()[1]+f.win.outerHeight<=1040);
});
test('valid negative work area: never clamp left/top to primary zero', () => {
    const f=fixture(); Object.assign(f.screen,{availLeft:-1920,availTop:-1080});
    f.place(-1800,-850); f.c.showRequest(2); f.advance();
    assert.ok(f.position()[0]<0 && f.position()[1]<0);
});
test('screen.left/top fallback is used only when avail origin is a missing number', () => {
    const f=fixture({availLeft:null,availTop:null,left:-1920,top:0});
    assert.equal(f.c.getCurrentWorkArea().left,-1920);
});
test('mixed-DPI screen dimensions changing in watcher preserve external location', () => {
    const f=fixture(); f.place(2300,160); f.c.startScreenWatcher();
    f.screen.availWidth=2560; f.screen.availHeight=1400; f.tickWatch();
    assert.deepEqual(f.position(),[2300,160]);
});
test('resize side effect restores original coordinates without a frame-offset drift', () => {
    const f=fixture(); f.place(-1850,150); f.win.resizeShift=[120,25];
    f.c.resizeWindowKeepingPosition(1750,700);
    assert.deepEqual(f.position(),[-1850,150]);
});
test('minimized window is not resized or moved by watcher/layout', () => {
    const f=fixture({availLeft:0,availTop:0}); f.place(-32000,-32000);
    const sizeCount=f.sizes.length; f.c.resizeApp(false); f.advance();
    assert.deepEqual(f.position(),[-32000,-32000]);
    assert.equal(f.sizes.length,sizeCount);
});
test('user drag between resize and delayed fit is respected', () => {
    const f=fixture({availLeft:0,availTop:0}); f.place(100,150);
    f.c.showRequest(2); f.advance(15); f.place(2050,150); f.advance();
    assert.deepEqual(f.position(),[2050,150]);
    assert.equal(f.moves.length,0);
});
test('superseded delayed fit does not run against a newer layout request', () => {
    const f=fixture({availLeft:0,availTop:0});
    let fits=0;
    const fit=f.c.fitCurrentWindowIntoWorkArea;
    f.c.fitCurrentWindowIntoWorkArea=function() {fits++; return fit.apply(this,arguments);};
    f.c.showRequest(2); f.advance(15); f.c.showRequest(3);
    const sizeCount=f.sizes.length, fitCount=fits; f.advance(10);
    assert.equal(f.sizes.length,sizeCount);
    assert.equal(fits,fitCount,'an obsolete layout correction ran');
    f.advance(); assert.equal(f.c.visibleRequestCount,3);
});
test('unknown window coordinates never cause an invented move to zero', () => {
    const f=fixture({availLeft:0,availTop:0});
    f.win.screenLeft=undefined; f.win.screenTop=undefined;
    f.moves.length=0; f.c.showRequest(2); f.advance();
    assert.equal(f.moves.length,0);
});

test('native layout route uses natural content dimensions without browser relocation', () => {
    const f=fixture(); f.place(2040,750);
    const requests=[];
    f.c.MentionLayout={request:(w,h,initial)=>{requests.push([w,h,initial]); return true;}};
    const count=f.sizes.length;
    f.c.showRequest(2); f.advance(); f.c.showRequest(3); f.advance();
    assert.equal(requests.length,2);
    assert.ok(requests[1][1]>requests[0][1]);
    assert.deepEqual(f.position(),[2040,750]);
    assert.equal(f.sizes.length,count,'browser resize must not race native monitor selection');
});

let failed=0;
for (const [name, fn] of cases) {
    try {fn(); console.log('PASS '+name);}
    catch (e) {failed++; console.error('FAIL '+name+'\n'+e.message);}
}
console.log(`${cases.length-failed}/${cases.length} passed`);
process.exitCode=failed ? 1 : 0;

// Browser DOM/layout test. Does not send requests or run Windows helpers.
const fs=require('node:fs'),path=require('node:path'),assert=require('node:assert/strict');
const {chromium}=require(process.env.CODEX_PRIMARY_RUNTIME_NODE_MODULES+'/playwright');
const folder=path.join(__dirname,'../メンション依頼フォーム_HTA_最新版');
const html=fs.readFileSync(path.join(folder,'メンション依頼フォーム.hta'),'utf8').replace(/^\uFEFF/,'')
    .replace(/<hta:application\b[\s\S]*?\/>/i,'').replace(/<script\b[^>]*>[\s\S]*?<\/script>/gi,'')
    .replace(/<link\b[^>]*>/gi,'').replace(' onload="initApp();"','');
(async()=>{
    const browser=await chromium.launch({executablePath:process.env.MENTION_TEST_BROWSER,headless:true});
    try{
        const page=await browser.newPage({viewport:{width:1540,height:744}});
        const errors=[];page.on('pageerror',e=>errors.push(e.message));
        await page.setContent(html);
        await page.addStyleTag({content:fs.readFileSync(path.join(folder,'mention-form.css'),'utf8').replace(/^\uFEFF/,'')});
        await page.addScriptTag({content:fs.readFileSync(path.join(folder,'mention-form.js'),'utf8')});
        await page.evaluate(()=>{
            populateTypes();populateBases();
            for(var n=1;n<=2;n++){
                for(var field of ['org','ca','mailMemo','processedAt']) $(field+n).value='test';
                $('attendanceWork'+n).checked=true;$('type'+n).selectedIndex=1;
            }
            // Force native-unavailable fallback with no usable monitor origin.
            getCurrentWorkArea=function(){return {left:null,top:null,width:1920,height:1040};};
            resizeApp(true);showRequest(2);showRequest(3);
        });
        await page.waitForFunction(()=>document.getElementById('app').style.visibility==='visible');
        for(const viewport of [{width:1540,height:744},{width:1280,height:600}]){
            await page.setViewportSize(viewport);
            await page.locator('#sendButton').scrollIntoViewIfNeeded();
            const box=await page.locator('#sendButton').boundingBox();
            assert.ok(box.y>=0 && box.y+box.height<=viewport.height,'send button is vertically reachable');
            assert.ok(box.x>=0 && box.x+box.width<=viewport.width,'send button is horizontally reachable');
            assert.ok(await page.evaluate(()=>window.scrollY>0),'document can scroll down');
            await page.locator('#request3 #dueDate3').fill('2026/09/15');
            assert.equal(await page.inputValue('#dueDate3'),'2026/09/15');
            await page.evaluate(()=>window.scrollTo(0,0));
            await page.locator('#org1').fill('edited');
            assert.equal(await page.inputValue('#org1'),'edited');
        }
        assert.deepEqual(errors,[]);
        console.log('Browser passed: three cards, vertical/horizontal scrolling, footer reachability, fields remain editable at 1540x744 and 1280x600.');
    }finally{await browser.close();}
})().catch(e=>{console.error(e);process.exitCode=1;});

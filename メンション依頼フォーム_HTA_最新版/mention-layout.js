// v28.31.2: asynchronous, per-form Windows layout helper. Geometry only.
var MentionLayout=(function(){
    var state=null, disabled=false, timer=null, serial=0;
    function arg(value){
        return '"'+String(value).replace(/(\\*)"/g,'$1$1\\"').replace(/(\\+)$/g,'$1$1')+'"';
    }
    function now(){ return new Date().getTime(); }
    function write(name,text){
        var file=state.fso.CreateTextFile(state.fso.BuildPath(state.folder,name),true,false);
        try{ file.Write(text); }finally{ file.Close(); }
    }
    function read(name){
        var path=state.fso.BuildPath(state.folder,name);
        if(!state.fso.FileExists(path)){ return ""; }
        var file=state.fso.OpenTextFile(path,1,false,0);
        try{ return file.ReadAll(); }finally{ file.Close(); }
    }
    function stop(){
        if(timer){ window.clearInterval(timer); timer=null; }
        if(state){ try{ write("stop.txt","STOP"); }catch(ignore){} }
    }
    function fail(){
        if(disabled){ return; }
        disabled=true; stop();
        if(typeof layoutAdjustmentUnavailable==="function"){ layoutAdjustmentUnavailable(); }
    }
    function dispatch(){
        if(!state.ready || state.active || !state.pending){ return; }
        var next=state.pending, view=getViewportSize(), outer=getWindowOuterSize(0,0);
        if(view.width<=0 || view.height<=0){ return; }
        var id=++serial;
        // Read viewport dimensions only at dispatch, after any previous native resize.
        write("request.txt",[id,next.width,next.height,Math.round(view.width),Math.round(view.height),
            Math.round(outer.width),Math.round(outer.height),next.initial?1:0,id].join("|"));
        state.pending=null; state.active=id; state.sent=now();
    }
    function poll(){
        if(disabled || !state){ return; }
        try{
            if(now()-state.beat>5000){ write("heartbeat.txt","OK"); state.beat=now(); }
            if(read("error.txt")){ fail(); return; }
            if(!state.ready){
                if(read("ready.txt")==="OK"){ state.ready=true; }
                else if(now()-state.started>5000){ fail(); return; }
            }
            if(state.active){
                var result=read("result.txt").split("|");
                if(Number(result[0])===state.active && (result[1]==="OK" || result[1]==="SKIP")){
                    state.active=0;
                }else if(now()-state.sent>10000){ fail(); return; }
            }
            dispatch();
        }catch(err){ fail(); }
    }
    function start(){
        try{
            var fso=new ActiveXObject("Scripting.FileSystemObject");
            var shell=new ActiveXObject("WScript.Shell");
            var base=getCurrentFolderPath();
            var script=fso.BuildPath(base,"mention-layout-host.ps1");
            var code=fso.BuildPath(base,"mention-layout-host.cs");
            var executable=fso.BuildPath(shell.ExpandEnvironmentStrings("%SystemRoot%"),"System32\\WindowsPowerShell\\v1.0\\powershell.exe");
            if(!fso.FileExists(script) || !fso.FileExists(code) || !fso.FileExists(executable)){ return false; }
            var token=now().toString(36)+"-"+Math.floor(Math.random()*0x7fffffff).toString(36);
            var folder=fso.BuildPath(String(fso.GetSpecialFolder(2)),"MentionLayout-"+token);
            fso.CreateFolder(folder);
            state={fso:fso,folder:folder,started:now(),beat:now(),ready:false,active:0,pending:null};
            write("heartbeat.txt","OK");
            shell.Run(arg(executable)+" -NoLogo -NoProfile -NonInteractive -WindowStyle Hidden -File "+
                arg(script)+" -StateFolder "+arg(folder)+" -WindowTitle "+arg(document.title),0,false);
            timer=window.setInterval(poll,100);
            return true;
        }catch(err){ return false; }
    }
    function request(width,height,initial){
        if(disabled){ return false; }
        if(!state && !start()){ fail(); return false; }
        if(disabled || !state){ return false; }
        state.pending={width:Math.ceil(width),height:Math.ceil(height),initial:!!initial};
        try{ dispatch(); }catch(err){ fail(); return false; }
        return true;
    }
    return {request:request,stop:stop};
}());

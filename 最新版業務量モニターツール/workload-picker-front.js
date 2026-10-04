/* Windows Script Host / JScript (ES3). Only activates the uniquely titled picker. */
(function () {
    var title, shell, i;
    if (WScript.Arguments.length !== 1) { WScript.Quit(2); return; }
    title = String(WScript.Arguments.Item(0));
    if (!/\[[a-z0-9]+-[0-9]+\]$/.test(title)) { WScript.Quit(2); return; }
    try { shell = new ActiveXObject("WScript.Shell"); }
    catch (unavailable) { WScript.Quit(3); return; }
    for (i = 0; i < 120; i++) {
        WScript.Sleep(100);
        try {
            if (shell.AppActivate(title)) { WScript.Quit(0); return; }
        } catch (notReady) {}
    }
    WScript.Quit(1);
}());

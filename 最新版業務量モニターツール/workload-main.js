/*
 * 業務量モニター v0.5 / CSV・Excel取込
 * ES3構文。設定は workload-config.js に集約。
 * parseCSV / buildData は外部I/Oを持たず、検証と実際の取込で共用する。
 * runWeekday は全データの検証完了後に結果を返す。画面の更新はUI側のみ。
 */
var WorkloadImport = (function () {
    var nativeExcel = null; // このフォーム専用に起動したExcelだけを保持する。
    var pickerSequence = 0;
    var pickerSession = new Date().getTime().toString(36) + Math.random().toString(36).slice(2, 8);
    function trim(value) {
        return String(value === null || typeof value === "undefined" ? "" : value).replace(/^\s+|\s+$/g, "");
    }
    function normalize(value) { return trim(value).replace(/[\s\u3000]+/g, ""); }
    function keyOf(organization, name) {
        var a = normalize(organization), b = normalize(name);
        return "@" + a.length + ":" + a + "|" + b.length + ":" + b;
    }
    function safeId(prefix, key) {
        var result = prefix, i, part;
        for (i = 0; i < key.length; i++) {
            part = key.charCodeAt(i).toString(16);
            result += ("0000" + part).slice(-4);
        }
        return result;
    }
    function columnIndex(column) {
        var s = trim(column).toUpperCase(), n = 0, i;
        if (!/^[A-Z]{1,3}$/.test(s)) { throw new Error("列の設定を確認してください：" + column); }
        for (i = 0; i < s.length; i++) { n = n * 26 + s.charCodeAt(i) - 64; }
        if (n > 16384) { throw new Error("Excelの列範囲を超えています：" + column); }
        return n - 1;
    }
    function columnLabel(index) {
        var n = index + 1, result = "";
        while (n > 0) { n--; result = String.fromCharCode(65 + n % 26) + result; n = Math.floor(n / 26); }
        return result;
    }
    function configOrDefault(config) {
        if (config) { return config; }
        if (typeof WorkloadConfig !== "undefined") { return WorkloadConfig; }
        throw new Error("workload-config.js が見つからないか、設定の書式が不正です。HTAと同じフォルダに配置してください。");
    }
    function validateConfig(config) {
        var c = configOrDefault(config), names = ["unread", "flag", "assignment"], i, j, section, cols, items, seen;
        if (!c.paths || !trim(c.paths.unreadRoot) || !trim(c.paths.flagRoot) || !trim(c.paths.assignmentFile)) {
            throw new Error("workload-config.js の paths に3つのパスを設定してください。");
        }
        for (i = 0; i < names.length; i++) {
            section = c[names[i]];
            if (!section || typeof section.startRow !== "number" || section.startRow < 2 || Math.floor(section.startRow) !== section.startRow) {
                throw new Error(names[i] + " の startRow は2以上の整数で指定してください。");
            }
            cols = section.columns || {};
            items = names[i] === "unread" ? [cols.organization, cols.caName, cols.unreadBase, cols.unsorted] :
                (names[i] === "flag" ? [cols.organization, cols.caName, cols.flag] :
                [cols.organization, cols.caName, cols.calib, cols.team, cols.staff, cols.canvas]);
            if (names[i] === "unread") {
                if (Object.prototype.toString.call(cols.unreadSubtract) !== "[object Array]") {
                    throw new Error("unreadSubtract は列記号の配列で指定してください。");
                }
                items = items.concat(cols.unreadSubtract);
            }
            seen = {};
            for (j = 0; j < items.length; j++) {
                var ix = columnIndex(items[j]);
                if (seen["@" + ix]) { throw new Error(names[i] + " の列指定が重複しています：" + items[j]); }
                seen["@" + ix] = true;
            }
            if (names[i] !== "assignment" && !/^(auto|utf-8|shift_jis|unicode|unicodeFFFE)$/i.test(trim(section.charset))) {
                throw new Error(names[i] + " の charset は auto / utf-8 / shift_jis / unicode などで指定してください。");
            }
        }
        if (!trim(c.assignment.sheetName)) { throw new Error("assignment.sheetName を指定してください。"); }
        if (!c.capacity || typeof c.capacity.hourlyRate !== "number" || !isFinite(c.capacity.hourlyRate) || c.capacity.hourlyRate <= 0 ||
            Object.prototype.toString.call(c.capacity.excludedStaffNames) !== "[object Array]" ||
            Object.prototype.toString.call(c.capacity.excludedTeams) !== "[object Array]") {
            throw new Error("capacity の処理件数・除外担当者名・除外チームを確認してください。");
        }
        if (!c.merge || !trim(c.merge.unassignedStaffName) || !trim(c.merge.unassignedTeam) ||
            !/^(sum|error)$/.test(c.merge.duplicateLogRows)) {
            throw new Error("merge の担当なし表示名・重複行の扱いを確認してください。");
        }
        return c;
    }
    function dateFolder(date, pattern) {
        var mm = ("0" + (date.getMonth() + 1)).slice(-2), dd = ("0" + date.getDate()).slice(-2);
        if (!/^(yyyymmdd|yyyymm|mmdd)$/.test(pattern)) {
            throw new Error("folderFormat は mmdd / yyyymm / yyyymmdd から指定してください。");
        }
        return pattern.replace("yyyy", String(date.getFullYear())).replace("mm", mm).replace("dd", dd);
    }
    function windowsPath(path) { return trim(path).replace(/\//g, "\\"); }
    function joinPath(root, child) { return windowsPath(root).replace(/\\+$/g, "") + "\\" + child; }
    function planPaths(date, config) {
        var c = validateConfig(config);
        return {
            unreadFolder: joinPath(c.paths.unreadRoot, dateFolder(date, c.unread.folderFormat)),
            flagFolder: joinPath(c.paths.flagRoot, dateFolder(date, c.flag.folderFormat)),
            assignmentFile: windowsPath(c.paths.assignmentFile)
        };
    }
    function parseCSV(text) {
        var source = String(text).replace(/^\uFEFF/, ""), rows = [], row = [], field = "", i = 0;
        var quoted = false, closed = false, ch, next;
        function finishField() { row.push(field); field = ""; closed = false; }
        function finishRow() { finishField(); rows.push(row); row = []; }
        for (i = 0; i < source.length; i++) {
            ch = source.charAt(i); next = source.charAt(i + 1);
            if (quoted) {
                if (ch === '"' && next === '"') { field += '"'; i++; }
                else if (ch === '"') { quoted = false; closed = true; }
                else { field += ch; }
            } else if (ch === ",") { finishField(); }
            else if (ch === "\r" || ch === "\n") {
                finishRow(); if (ch === "\r" && next === "\n") { i++; }
            } else if (ch === '"') {
                if (closed || trim(field)) { throw new Error("CSV " + (rows.length + 1) + "行目の引用符を確認してください。"); }
                field = ""; quoted = true;
            } else if (closed) {
                if (!/\s/.test(ch)) { throw new Error("CSV " + (rows.length + 1) + "行目の閉じ引用符の後に不正な文字があります。"); }
            } else { field += ch; }
        }
        if (quoted) { throw new Error("CSV " + (rows.length + 1) + "行目の引用符が閉じていません。"); }
        if (row.length || field || closed) { finishRow(); }
        return rows;
    }
    function blankRow(row) {
        var i; for (i = 0; i < row.length; i++) { if (trim(row[i])) { return false; } }
        return true;
    }
    function countValue(value, context) {
        var s = trim(value).replace(/[０-９]/g, function (ch) { return String.fromCharCode(ch.charCodeAt(0) - 65248); });
        if (!s) { return 0; }
        s = s.replace(/[\s\u3000]/g, "");
        if (!/^(?:\d+|\d{1,3}(?:,\d{3})+)$/.test(s)) {
            throw new Error(context + "：件数が非負の整数ではありません（" + trim(value) + "）。");
        }
        var valueNumber = Number(s.replace(/,/g, ""));
        if (!isFinite(valueNumber) || valueNumber > 9007199254740991) { throw new Error(context + "：件数が大きすぎます。"); }
        return valueNumber;
    }
    function rowIdentity(row, columns, context) {
        var organization = trim(row[columnIndex(columns.organization)]), name = trim(row[columnIndex(columns.caName)]);
        if (!organization || !name) { throw new Error(context + "：組織またはCA名が空欄です。指定列を確認してください。"); }
        if (/\uFFFD/.test(organization + name)) { throw new Error(context + "：文字化けを検出しました。設定の charset を確認してください。"); }
        return { key: keyOf(organization, name), organization: organization, name: name };
    }
    function requireColumns(row, columns, context) {
        var i, maximum = 0;
        for (i = 0; i < columns.length; i++) { maximum = Math.max(maximum, columnIndex(columns[i])); }
        if (row.length <= maximum) { throw new Error(context + "：" + columnLabel(maximum) + "列まで必要ですが、列数が足りません。"); }
    }
    function readLogRows(rows, kind, config, label) {
        var c = config[kind], cols = c.columns, result = { map: {}, order: [], rowCount: 0, duplicateCount: 0 };
        var required = kind === "unread" ? [cols.organization, cols.caName, cols.unreadBase, cols.unsorted].concat(cols.unreadSubtract) : [cols.organization, cols.caName, cols.flag];
        var i, j, row, context, identity, current, unread, flag, unsorted;
        if (rows.length < c.startRow - 1) { throw new Error(label + "：見出し行がありません。"); }
        requireColumns(rows[c.startRow - 2], required, label + " 見出し");
        for (i = c.startRow - 1; i < rows.length; i++) {
            row = rows[i]; if (blankRow(row)) { continue; }
            context = label + " " + (i + 1) + "行目";
            requireColumns(row, required, context); identity = rowIdentity(row, cols, context);
            unread = 0; flag = 0; unsorted = 0;
            if (kind === "unread") {
                unread = countValue(row[columnIndex(cols.unreadBase)], context + " " + cols.unreadBase + "列");
                for (j = 0; j < cols.unreadSubtract.length; j++) {
                    unread -= countValue(row[columnIndex(cols.unreadSubtract[j])], context + " " + cols.unreadSubtract[j] + "列");
                }
                if (unread < 0) { throw new Error(context + "：H相当列から控除列を引いた純粋未読が負の値です（" + unread + "）。元データと列設定を確認してください。"); }
                unsorted = countValue(row[columnIndex(cols.unsorted)], context + " " + cols.unsorted + "列");
            } else { flag = countValue(row[columnIndex(cols.flag)], context + " " + cols.flag + "列"); }
            current = result.map[identity.key];
            if (!current) {
                current = { key: identity.key, organization: identity.organization, name: identity.name, unread: 0, flag: 0, unsorted: 0 };
                result.map[identity.key] = current; result.order.push(identity.key);
            } else {
                if (config.merge.duplicateLogRows === "error") { throw new Error(context + "：同じ組織・CA名が複数行あります。"); }
                result.duplicateCount++;
            }
            current.unread += unread; current.flag += flag; current.unsorted += unsorted; result.rowCount++;
        }
        return result;
    }
    function listContains(list, value) {
        var i; for (i = 0; i < list.length; i++) { if (normalize(list[i]) === normalize(value)) { return true; } }
        return false;
    }
    function validURL(value) {
        var s = trim(value);
        return /^https?:\/\/[^\s\/?#]+(?:[\/?#][^\s]*)?$/i.test(s) && !/[\x00-\x1f\x7f"<>]/.test(s) ? s : "";
    }
    function buildData(unreadRows, flagRows, assignmentRows, previousStaff, config, linksByRow) {
        var c = validateConfig(config), unread = readLogRows(unreadRows, "unread", c, "未読CSV");
        var flag = readLogRows(flagRows, "flag", c, "フラグCSV");
        return buildFromLogs(unread, flag, assignmentRows, previousStaff, c, linksByRow);
    }
    function buildFromLogs(unread, flag, assignmentRows, previousStaff, c, linksByRow) {
        var roster = {}, all = {}, order = [], supports = {};
        var cols = c.assignment.columns, required = [cols.organization, cols.caName, cols.calib, cols.team, cols.staff, cols.canvas];
        var i, j, row, identity, item, old, key, team, staffName, staffKey, staff, ca, linkInfo, rawCalib, rawCanvas;
        var staffMap = {}, staffList = [], stats = { unreadRows: unread.rowCount, flagRows: flag.rowCount, assignmentRows: 0,
            assignmentDuplicates: 0, unreadDuplicates: unread.duplicateCount, flagDuplicates: flag.duplicateCount,
            unmatchedCA: 0, noUnreadCA: 0, noFlagCA: 0, invalidLinks: 0, caCount: 0 };
        function remember(k) { if (!all[k]) { all[k] = true; order.push(k); } }
        if (assignmentRows.length < c.assignment.startRow - 1) { throw new Error("振分表：見出し行がありません。"); }
        requireColumns(assignmentRows[c.assignment.startRow - 2], required, "振分表 見出し");
        for (i = c.assignment.startRow - 1; i < assignmentRows.length; i++) {
            row = assignmentRows[i]; if (blankRow(row)) { continue; }
            requireColumns(row, required, "振分表 " + (i + 1) + "行目");
            identity = rowIdentity(row, cols, "振分表 " + (i + 1) + "行目");
            staffName = trim(row[columnIndex(cols.staff)]) || c.merge.unassignedStaffName;
            team = trim(row[columnIndex(cols.team)]) || (staffName === c.merge.unassignedStaffName ? c.merge.unassignedTeam : "未設定");
            linkInfo = linksByRow && linksByRow["@" + (i + 1)] || {};
            rawCalib = trim(linkInfo.calib || row[columnIndex(cols.calib)]);
            rawCanvas = trim(linkInfo.canvas || row[columnIndex(cols.canvas)]);
            item = { key: identity.key, organization: identity.organization, name: identity.name, staff: staffName, team: team,
                calibUrl: validURL(rawCalib), canvasUrl: validURL(rawCanvas) };
            if (rawCalib && !item.calibUrl) { stats.invalidLinks++; }
            if (rawCanvas && !item.canvasUrl) { stats.invalidLinks++; }
            old = roster[identity.key];
            if (old && (normalize(old.staff) !== normalize(item.staff) || normalize(old.team) !== normalize(item.team) || old.calibUrl !== item.calibUrl || old.canvasUrl !== item.canvasUrl)) {
                throw new Error("振分表 " + (i + 1) + "行目：" + identity.organization + " / " + identity.name + " の担当者・チーム・リンクが別の行と一致しません。");
            }
            if (old) { stats.assignmentDuplicates++; }
            else { roster[identity.key] = item; remember(identity.key); }
            stats.assignmentRows++;
        }
        if (!stats.assignmentRows) { throw new Error("「" + c.assignment.sheetName + "」に有効な組織・CA名の行がありません。シート名・開始行・列を確認してください。"); }
        for (i = 0; i < unread.order.length; i++) { remember(unread.order[i]); }
        for (i = 0; i < flag.order.length; i++) { remember(flag.order[i]); }
        previousStaff = previousStaff || [];
        for (i = 0; i < previousStaff.length; i++) {
            for (j = 0; j < previousStaff[i].cas.length; j++) {
                old = previousStaff[i].cas[j];
                if (old.importKey) { supports[old.importKey] = old.supports.slice(0, 3); }
            }
        }
        for (i = 0; i < order.length; i++) {
            key = order[i]; item = roster[key];
            if (!item) {
                identity = unread.map[key] || flag.map[key];
                item = { key: key, organization: identity.organization, name: identity.name, staff: c.merge.unassignedStaffName,
                    team: c.merge.unassignedTeam, calibUrl: "", canvasUrl: "" };
                stats.unmatchedCA++;
            }
            staffKey = "@" + normalize(item.staff);
            // 担当なし等の区分は、異なるチームに現れても別の表示行として扱う。
            if (listContains(c.capacity.excludedStaffNames, item.staff) || item.staff === c.merge.unassignedStaffName) {
                staffKey += "|" + normalize(item.team);
            }
            staff = staffMap[staffKey];
            if (staff && normalize(staff.team) !== normalize(item.team)) {
                throw new Error("振分表：担当者「" + item.staff + "」に複数のチームが設定されています（" + staff.team + " / " + item.team + "）。チーム欄を確認してください。");
            }
            if (!staff) {
                staff = { id: safeId("staff", staffKey), name: item.staff, team: item.team, assignmentType: "通常担当", cas: [] };
                if (listContains(c.capacity.excludedStaffNames, item.staff) || item.staff === c.merge.unassignedStaffName) { staff.assignmentType = "担当なし"; }
                else if (listContains(c.capacity.excludedTeams, item.team)) { staff.assignmentType = "担当なし"; }
                staffMap[staffKey] = staff; staffList.push(staff);
            }
            ca = { id: safeId("ca", key), importKey: key, organization: item.organization, name: item.name,
                unsorted: unread.map[key] ? unread.map[key].unsorted : 0,
                unread: unread.map[key] ? unread.map[key].unread : 0,
                flag: flag.map[key] ? flag.map[key].flag : 0,
                calibUrl: item.calibUrl, canvasUrl: item.canvasUrl, supports: supports[key] || [] };
            if (!unread.map[key]) { stats.noUnreadCA++; }
            if (!flag.map[key]) { stats.noFlagCA++; }
            staff.cas.push(ca); stats.caCount++;
        }
        return { staffData: staffList, stats: stats };
    }
    function detectCharset(bytes) {
        var i = 0, b, b2, remaining, count, j;
        if (bytes.length >= 3 && bytes[0] === 239 && bytes[1] === 187 && bytes[2] === 191) { return "utf-8"; }
        if (bytes.length >= 2 && bytes[0] === 255 && bytes[1] === 254) { return "unicode"; }
        if (bytes.length >= 2 && bytes[0] === 254 && bytes[1] === 255) { return "unicodeFFFE"; }
        while (i < bytes.length) {
            b = bytes[i++]; if (b < 128) { continue; }
            if (b >= 194 && b <= 223) { count = 1; }
            else if (b >= 224 && b <= 239) { count = 2; }
            else if (b >= 240 && b <= 244) { count = 3; }
            else { return "shift_jis"; }
            remaining = bytes.length - i; if (remaining < count) { return "shift_jis"; }
            b2 = bytes[i];
            if ((b === 224 && b2 < 160) || (b === 237 && b2 >= 160) || (b === 240 && b2 < 144) || (b === 244 && b2 >= 144)) { return "shift_jis"; }
            for (j = 0; j < count; j++) { if (bytes[i] < 128 || bytes[i] > 191) { return "shift_jis"; } i++; }
        }
        return "utf-8";
    }
    function isHTA() { return /\.hta$/i.test(String(window.location.pathname)) && typeof ActiveXObject !== "undefined"; }
    function readNativeCSV(path, charset) {
        var stream = null, bytes = null, text, usedCharset;
        try {
            stream = new ActiveXObject("ADODB.Stream"); stream.Type = 1; stream.Open(); stream.LoadFromFile(path);
            usedCharset = charset;
            if (String(charset).toLowerCase() === "auto") {
                bytes = new VBArray(stream.Read()).toArray(); usedCharset = detectCharset(bytes);
            }
            bytes = null; stream.Position = 0; stream.Type = 2; stream.Charset = usedCharset; text = stream.ReadText(-1);
            if (/\uFFFD/.test(text)) { throw new Error("文字化けを検出しました。workload-config.js の charset を utf-8 または shift_jis に指定してお試しください。"); }
            return { rows: parseCSV(text), charset: usedCharset };
        } catch (e) { throw new Error("CSVを読み込めません：" + path + "\n" + (e.message || e.description || String(e))); }
        finally { if (stream) { try { stream.Close(); } catch (ignore) {} } stream = null; bytes = null; }
    }
    function pickerFolder(fso, folder, parent) {
        if (fso.FolderExists(folder)) { return { path: folder, allowSubfolders: false }; }
        var root = windowsPath(parent);
        if (!fso.FolderExists(root)) { throw new Error("CSVを選択する親フォルダが見つかりません：\n" + root); }
        return { path: root, allowSubfolders: true };
    }
    function startPickerForeground(fso, title) {
        var path = String(window.location.pathname), host = String(window.location.hostname || ""), helper, shell = null, command;
        try { path = decodeURIComponent(path); } catch (rawPath) {}
        path = windowsPath(path);
        if (host && host.toLowerCase() !== "localhost" && !/^\\\\/.test(path)) { path = "\\\\" + host + "\\" + path.replace(/^\\+/, ""); }
        path = path.replace(/^\\([a-z]:\\)/i, "$1");
        helper = joinPath(fso.GetParentFolderName(path), "workload-picker-front.js");
        if (!fso.FileExists(helper)) { throw new Error("CSV選択画面の前面化ファイルがありません。workload-picker-front.jsをHTAと同じフォルダに配置してください。"); }
        if (/["\r\n]/.test(helper + title)) { throw new Error("CSV選択画面のパスまたはタイトルに使用できない文字があります。"); }
        try {
            shell = new ActiveXObject("WScript.Shell");
            command = '"' + shell.ExpandEnvironmentStrings("%SystemRoot%\\System32\\wscript.exe") + '" //B //Nologo //E:JScript //T:15 "' + helper + '" "' + title + '"';
            shell.Run(command, 0, false);
        } catch (launchError) {
            throw new Error("CSV選択画面を前面にする補助処理を起動できません。Windows Script Hostの利用可否を確認してください。\n" + (launchError.message || launchError.description || String(launchError)));
        } finally { shell = null; }
    }
    function pickCSV(excel, fso, folder, title, allowSubfolders) {
        var dialog = null, selected = "", selectedFolder, baseFolder;
        var dialogTitle = title + " [" + pickerSession + "-" + (++pickerSequence) + "]";
        try {
            /* Excel本体は非表示。ShowでHTAが待機している間、補助処理がこの選択画面だけを前面化する。 */
            excel.Visible = false;
            dialog = excel.FileDialog(3); // msoFileDialogFilePicker
            dialog.Title = dialogTitle; dialog.AllowMultiSelect = false; dialog.Filters.Clear(); dialog.Filters.Add("CSVファイル", "*.csv");
            dialog.InitialFileName = folder.replace(/\\+$/g, "") + "\\";
            startPickerForeground(fso, dialogTitle);
            if (dialog.Show() !== -1) { return ""; }
            selected = String(dialog.SelectedItems.Item(1));
            if (!/\.csv$/i.test(selected) || !fso.FileExists(selected)) { throw new Error("存在するCSVファイルを選択してください。"); }
            selectedFolder = windowsPath(fso.GetAbsolutePathName(fso.GetParentFolderName(selected))).replace(/\\+$/g, "").toLowerCase();
            baseFolder = windowsPath(fso.GetAbsolutePathName(folder)).replace(/\\+$/g, "").toLowerCase();
            if (selectedFolder !== baseFolder && !(allowSubfolders && selectedFolder.indexOf(baseFolder + "\\") === 0)) {
                throw new Error("指定のフォルダ内のCSVを選択してください：\n" + folder);
            }
            return selected;
        } finally {
            dialog = null;
            try { excel.Visible = false; } catch (hidePickerOwner) {}
        }
    }
    function arrayRows(value) {
        var array = new VBArray(value), rows = [], r, col, row, rl = array.lbound(1), ru = array.ubound(1), cl = array.lbound(2), cu = array.ubound(2);
        for (r = rl; r <= ru; r++) {
            row = [];
            for (col = cl; col <= cu; col++) {
                var v = array.getItem(r, col);
                row.push(v === null || typeof v === "undefined" ? "" : String(v));
            }
            rows.push(row);
        }
        return rows;
    }
    function readAssignment(excel, path, config, progress) {
        var book = null, sheet = null, range = null, linksRange = null, hyperlinks = null, link = null, formulas = null;
        var rows, links = {}, section = config.assignment, cols = section.columns, maximum = 0, lastRow, i, name, n, r, url, formula;
        var linkColumns = ["calib", "canvas"], property, linkCount, linkIndex, caIndex = columnIndex(cols.caName);
        var security = excel.AutomationSecurity, screenUpdating = excel.ScreenUpdating;
        function report(message) { if (progress) { progress(message); } }
        try {
            excel.ScreenUpdating = false;
            excel.AutomationSecurity = 3; // msoAutomationSecurityForceDisable
            try { book = excel.Workbooks.Open(path, 0, true); } // 外部リンクを更新せず、読み取り専用
            finally { excel.AutomationSecurity = security; }
            try { sheet = book.Worksheets.Item(section.sheetName); }
            catch (notFound) { throw new Error("振分表に「" + section.sheetName + "」シートが見つかりません。設定の sheetName を確認してください。"); }
            for (name in cols) { if (Object.prototype.hasOwnProperty.call(cols, name)) { maximum = Math.max(maximum, columnIndex(cols[name])); } }
            lastRow = Math.max(Number(sheet.Cells(sheet.Rows.Count, columnIndex(cols.organization) + 1).End(-4162).Row),
                Number(sheet.Cells(sheet.Rows.Count, columnIndex(cols.caName) + 1).End(-4162).Row)); // xlUp
            if (lastRow < section.startRow) { throw new Error("「" + section.sheetName + "」に取込対象のCAがありません。"); }
            report("3/3 振分表の値をまとめて読み込んでいます。");
            range = sheet.Range("A1:" + columnLabel(maximum) + lastRow); rows = arrayRows(range.Value2); range = null;
            for (i = 0; i < linkColumns.length; i++) {
                property = linkColumns[i]; linkIndex = columnIndex(cols[property]); formulas = null;
                report("3/3 " + (property === "calib" ? "CALIB" : "Canvas") + "リンクを読み込んでいます。");
                linksRange = sheet.Range(cols[property] + section.startRow + ":" + cols[property] + lastRow);
                hyperlinks = linksRange.Hyperlinks; linkCount = hyperlinks.Count;
                for (n = 1; n <= linkCount; n++) {
                    link = hyperlinks.Item(n); r = Number(link.Range.Row); url = trim(link.Address);
                    if (url) { if (!links["@" + r]) { links["@" + r] = {}; } links["@" + r][property] = url; }
                    link = null;
                }
                hyperlinks = null;
                // 数式は列単位で一括取得し、行ごとのExcel COM呼出しを避ける。
                // 直接URL／ハイパーリンクで全行を解決できる列は、数式取得自体を省略する。
                for (r = section.startRow; r <= lastRow; r++) {
                    if (validURL(rows[r - 1][linkIndex]) || links["@" + r] && links["@" + r][property]) { continue; }
                    if (!trim(rows[r - 1][caIndex])) { continue; }
                    if (!formulas) {
                        // Excelは1セルだけの範囲では配列ではなく単一値を返す。
                        formulas = lastRow === section.startRow ? [[linksRange.Formula]] : arrayRows(linksRange.Formula);
                    }
                    formula = String(formulas[r - section.startRow][0]);
                    var match = /^=\s*HYPERLINK\s*\(\s*"((?:[^"]|"")*)"\s*[,;)]/i.exec(formula);
                    if (match) {
                        if (!links["@" + r]) { links["@" + r] = {}; }
                        links["@" + r][property] = match[1].replace(/""/g, '"');
                    }
                }
                formulas = null; linksRange = null;
            }
            return { rows: rows, linksByRow: links };
        } finally {
            formulas = null; link = null; hyperlinks = null; linksRange = null; range = null; sheet = null;
            if (book) { try { book.Close(false); } catch (ignore) {} } book = null;
            try { excel.ScreenUpdating = screenUpdating; } catch (restoreScreen) {}
        }
    }
    function releaseNativeExcel() {
        var owned = nativeExcel;
        nativeExcel = null;
        if (owned) { try { owned.Quit(); } catch (ignore) {} }
        owned = null;
        if (typeof CollectGarbage === "function") { CollectGarbage(); }
    }
    function acquireNativeExcel() {
        var candidate = null;
        if (nativeExcel) {
            try { if (nativeExcel.Workbooks.Count === 0) { return nativeExcel; } }
            catch (staleExcel) {}
            releaseNativeExcel();
        }
        try {
            candidate = new ActiveXObject("Excel.Application");
            candidate.Visible = false; candidate.DisplayAlerts = false; candidate.EnableEvents = false; candidate.AskToUpdateLinks = false;
            nativeExcel = candidate;
            return nativeExcel;
        } catch (excelError) {
            if (candidate) { try { candidate.Quit(); } catch (ignore) {} }
            candidate = null;
            throw new Error("Windows版Excelを起動できません。ExcelのインストールとHTAからの利用可否を確認してください。\n" + (excelError.message || excelError.description || ""));
        }
    }
    function prepareNativeExcel() {
        if (!isHTA()) { return false; }
        try { acquireNativeExcel(); return true; }
        catch (unavailable) { return false; } // エラーの案内はログ集計押下時に行う。
    }
    function runWeekday(previousStaff, progress) {
        var config, paths, excel = null, fso = null, unreadFile, flagFile, unread, flag, roster, result, unreadPicker, flagPicker;
        var unreadLog, flagLog;
        function report(message) { if (progress) { progress(message); } }
        try {
            config = validateConfig(); paths = planPaths(new Date(), config);
            if (!isHTA()) { throw new Error("ログ集計はWindowsで WorkloadRanking.hta を開いて実行してください。ブラウザ用ファイルは画面確認用です。"); }
            report("CSVフォルダ・振分表の場所を確認しています。");
            fso = new ActiveXObject("Scripting.FileSystemObject");
            unreadPicker = pickerFolder(fso, paths.unreadFolder, config.paths.unreadRoot);
            flagPicker = pickerFolder(fso, paths.flagFolder, config.paths.flagRoot);
            if (!fso.FileExists(paths.assignmentFile)) { throw new Error("振分表が見つかりません：\n" + paths.assignmentFile); }
            report("CSV選択用Excelを準備しています。"); excel = acquireNativeExcel();
            report("1/3 未読CSVを選択してください。");
            unreadFile = pickCSV(excel, fso, unreadPicker.path, unreadPicker.allowSubfolders ? "① 未読CSVを選択（親フォルダ）" : "① 当日の未読CSVを選択", unreadPicker.allowSubfolders);
            if (!unreadFile) { return { cancelled: true }; }
            report("未読CSVを読み込んでいます。"); unread = readNativeCSV(unreadFile, config.unread.charset);
            unreadLog = readLogRows(unread.rows, "unread", config, "未読CSV");
            report("2/3 フラグCSVを選択してください。");
            flagFile = pickCSV(excel, fso, flagPicker.path, flagPicker.allowSubfolders ? "② フラグCSVを選択（親フォルダ）" : "② 当月のフラグCSVを選択", flagPicker.allowSubfolders);
            if (!flagFile) { return { cancelled: true }; }
            report("フラグCSVを読み込んでいます。"); flag = readNativeCSV(flagFile, config.flag.charset);
            flagLog = readLogRows(flag.rows, "flag", config, "フラグCSV");
            report("3/3 「" + config.assignment.sheetName + "」を読み込んでいます。");
            roster = readAssignment(excel, paths.assignmentFile, config, report);
            report("CSVと振分表を照合しています。");
            // CSV選択直後の検証・集計結果を使う。同じCSVを再集計しない。
            result = buildFromLogs(unreadLog, flagLog, roster.rows, previousStaff, config, roster.linksByRow);
            result.sources = { unreadFile: unreadFile, flagFile: flagFile, assignmentFile: paths.assignmentFile,
                sheetName: config.assignment.sheetName, unreadCharset: unread.charset, flagCharset: flag.charset };
            result.importedAt = new Date(); result.hourlyRate = config.capacity.hourlyRate;
            return result;
        } finally {
            /* 配列の結果を残し、完了・キャンセル・エラー時に専用Excelも終了する。 */
            excel = null; fso = null;
            releaseNativeExcel();
            try { window.focus(); } catch (focusError) {}
        }
    }
    function openLink(url) {
        var target = validURL(url), shell = null;
        if (!target) { throw new Error("リンクが未設定か、http/httpsのURLではありません。"); }
        if (isHTA()) {
            try { shell = new ActiveXObject("Shell.Application"); shell.ShellExecute(target, "", "", "open", 1); }
            finally { shell = null; }
        } else { window.open(target, "_blank"); }
    }
    return { parseCSV: parseCSV, buildData: buildData, planPaths: planPaths, validateConfig: validateConfig,
        keyOf: keyOf, columnIndex: columnIndex, detectCharset: detectCharset, validURL: validURL,
        runWeekday: runWeekday, openLink: openLink,
        prepareNativeExcel: prepareNativeExcel, releaseNativeExcel: releaseNativeExcel };
}());

/*
 * 日付別の共有履歴。ES3 / Windows HTA。
 * 同じ共有フォルダの全フォームでWORKLOAD_LOCKを使い、再読込してから追記する。
 * 共有データはUTF-8 JSON。旧版の日付JSはデータだけを解析し、コードは実行しない。
 */
var WorkloadShare = (function () {
    var listener = null, pollTimer = null, retryTimer = null, stopped = true;
    var queue = [], current = null, knownDay = "", lastSignature = "", lastState = null;
    var sequence = 0, session = "w" + new Date().getTime().toString(36) + Math.random().toString(36).slice(2) + Math.random().toString(36).slice(2);
    var lockName = "WORKLOAD_LOCK", retryLimit = 50;
    function trim(value) { return String(value === null || typeof value === "undefined" ? "" : value).replace(/^\s+|\s+$/g, ""); }
    function normalized(value) { return trim(value).toLowerCase().replace(/[\s\u3000]+/g, ""); }
    function keyOf(organization, caName) {
        var a = trim(organization).replace(/[\s\u3000]+/g, ""), b = trim(caName).replace(/[\s\u3000]+/g, "");
        return "@" + a.length + ":" + a + "|" + b.length + ":" + b;
    }
    function pad(value) { return value < 10 ? "0" + value : String(value); }
    function dayKey(date) { return String(date.getFullYear()) + pad(date.getMonth() + 1) + pad(date.getDate()); }
    function timeText(date) { return date.getFullYear() + "-" + pad(date.getMonth() + 1) + "-" + pad(date.getDate()) + " " + pad(date.getHours()) + ":" + pad(date.getMinutes()) + ":" + pad(date.getSeconds()); }
    function json(value) { return JSON.stringify(value).replace(/\u2028/g, "\\u2028").replace(/\u2029/g, "\\u2029"); }
    function own(object, name) { return Object.prototype.hasOwnProperty.call(object, name); }
    function empty(day) { return { schemaVersion: 2, date: day, supports: [], commits: [], revision: 0, ids: {} }; }
    function legacyHeader(day) {
        return 'var WorkloadDaily = ' + json({ schemaVersion: 1, date: day, supports: [], commits: [] }) + ';\n' +
            'function WorkloadDailyAppend(record) {\n' +
            '    var i;\n' +
            '    for (i = 0; i < record.supports.length; i++) { WorkloadDaily.supports.push(record.supports[i]); }\n' +
            '    for (i = 0; i < record.commits.length; i++) { WorkloadDaily.commits.push(record.commits[i]); }\n' +
            '}\n';
    }
    function validateIdentity(record) {
        if (!record || typeof record.organization !== "string" || !trim(record.organization) || typeof record.caName !== "string" || !trim(record.caName)) {
            throw new Error("共有履歴の組織・CA名が不正です。");
        }
    }
    function validateTime(value) { if (typeof value !== "string" || !/^\d{4}-\d\d-\d\d \d\d:\d\d:\d\d$/.test(value)) { throw new Error("共有履歴の時刻が不正です。"); } }
    function committed(state, key) {
        var i;
        for (i = 0; i < state.commits.length; i++) { if (keyOf(state.commits[i].organization, state.commits[i].caName) === key) { return state.commits[i]; } }
        return null;
    }
    function active(state, key) {
        var result = [], i, row;
        for (i = 0; i < state.supports.length; i++) {
            row = state.supports[i];
            if (!row.completedAt && keyOf(row.organization, row.caName) === key) { result.push(row); }
        }
        return result;
    }
    function apply(state, record) {
        var i, j, item, found, key, list;
        if (!record || typeof record.id !== "string" || !/^[a-z0-9-]+$/i.test(record.id) || !record.supports || !record.commits ||
                Object.prototype.toString.call(record.supports) !== "[object Array]" || Object.prototype.toString.call(record.commits) !== "[object Array]") { throw new Error("共有履歴の行が不正です。"); }
        if (own(state.ids, "@" + record.id)) { return; }
        for (i = 0; i < record.supports.length; i++) {
            item = record.supports[i]; validateIdentity(item); validateTime(item.startedAt);
            if (typeof item.id !== "string" || !item.id || typeof item.supportName !== "string" || !trim(item.supportName) || typeof item.completedAt !== "string") { throw new Error("共有履歴のサポート情報が不正です。"); }
            key = keyOf(item.organization, item.caName); found = null;
            for (j = 0; j < state.supports.length; j++) { if (state.supports[j].id === item.id) { found = state.supports[j]; break; } }
            if (item.completedAt) {
                validateTime(item.completedAt);
                if (!found || keyOf(found.organization, found.caName) !== key || found.supportName !== item.supportName || found.startedAt !== item.startedAt) { throw new Error("共有履歴のサポート完了に対応する投入がありません。"); }
                if (!found.completedAt) { found.completedAt = item.completedAt; }
            } else {
                if (found || committed(state, key)) { throw new Error("共有履歴に重複した投入があります。"); }
                list = active(state, key);
                if (list.length >= 3) { throw new Error("共有履歴のサポート人数が3名を超えています。"); }
                for (j = 0; j < list.length; j++) { if (normalized(list[j].supportName) === normalized(item.supportName)) { throw new Error("共有履歴に同じ氏名のサポートが重複しています。"); } }
                state.supports.push({ id: item.id, organization: item.organization, caName: item.caName, supportName: item.supportName, startedAt: item.startedAt, completedAt: "" });
            }
        }
        for (i = 0; i < record.commits.length; i++) {
            item = record.commits[i]; validateIdentity(item); validateTime(item.committedAt);
            key = keyOf(item.organization, item.caName);
            if (committed(state, key) || active(state, key).length) { throw new Error("共有履歴のコミット情報が不正です。"); }
            state.commits.push({ organization: item.organization, caName: item.caName, committedAt: item.committedAt });
        }
        if (record.cancelCommits) {
            if (Object.prototype.toString.call(record.cancelCommits) !== "[object Array]") { throw new Error("共有履歴のコミット取消情報が不正です。"); }
            for (i = 0; i < record.cancelCommits.length; i++) {
                item = record.cancelCommits[i]; validateIdentity(item); validateTime(item.committedAt);
                key = keyOf(item.organization, item.caName); found = committed(state, key);
                if (!found || found.committedAt !== item.committedAt) { throw new Error("取り消すコミットが共有履歴と一致しません。"); }
                for (j = 0; j < state.commits.length; j++) { if (state.commits[j] === found) { state.commits.splice(j, 1); break; } }
            }
        }
        state.ids["@" + record.id] = true; state.revision++;
    }
    function parseLegacy(text, day) {
        var raw = String(text).replace(/^\uFEFF/, "").replace(/\r\n/g, "\n"), prefix = legacyHeader(day), state = empty(day), lines, i, match;
        if (raw.slice(0, prefix.length) !== prefix) { throw new Error("共有JSの形式が異なります。既存ファイルは上書きしません。"); }
        lines = raw.slice(prefix.length).split("\n");
        for (i = 0; i < lines.length; i++) {
            if (!trim(lines[i])) { continue; }
            match = /^WorkloadDailyAppend\((.*)\);$/.exec(lines[i]);
            if (!match) { throw new Error("共有JSの末尾または履歴行が不正です。既存ファイルは上書きしません。"); }
            apply(state, JSON.parse(match[1]));
        }
        return state;
    }
    function encode(state) {
        var ids = [], id;
        for (id in state.ids) { if (own(state.ids, id)) { ids.push(id.slice(1)); } }
        return JSON.stringify({ schemaVersion: 2, date: state.date, supports: state.supports, commits: state.commits,
            revision: state.revision, operations: ids }, null, 2) + "\n";
    }
    function parse(text, day) {
        var value, state = empty(day), i, row, ids = {}, id, arrayTag = "[object Array]";
        try { value = JSON.parse(String(text).replace(/^\uFEFF/, "")); }
        catch (badJSON) { throw new Error("共有JSONの形式が不正です。既存ファイルは上書きしません。\n" + (badJSON.message || String(badJSON))); }
        if (!value || value.schemaVersion !== 2 || value.date !== day || Object.prototype.toString.call(value.supports) !== arrayTag ||
                Object.prototype.toString.call(value.commits) !== arrayTag || Object.prototype.toString.call(value.operations) !== arrayTag ||
                typeof value.revision !== "number" || value.revision !== value.operations.length) { throw new Error("共有JSONの項目・日付・履歴番号が不正です。既存ファイルは上書きしません。"); }
        for (i = 0; i < value.operations.length; i++) {
            id = value.operations[i];
            if (typeof id !== "string" || !/^[a-z0-9-]+$/i.test(id) || own(ids, "@" + id)) { throw new Error("共有JSONの操作履歴が不正です。"); }
            ids["@" + id] = true;
        }
        for (i = 0; i < value.supports.length; i++) {
            row = value.supports[i];
            if (!row || typeof row.completedAt !== "string" || typeof row.id !== "string" || !/^[a-z0-9-]+$/i.test(row.id) || !own(ids, "@" + row.id)) { throw new Error("共有JSONのサポート履歴が不正です。"); }
            apply(state, { id: "s-" + i, supports: [{ id: row.id, organization: row.organization, caName: row.caName,
                supportName: row.supportName, startedAt: row.startedAt, completedAt: "" }], commits: [] });
            if (row.completedAt) { apply(state, { id: "f-" + i, supports: [row], commits: [] }); }
        }
        for (i = 0; i < value.commits.length; i++) { apply(state, { id: "c-" + i, supports: [], commits: [value.commits[i]] }); }
        state.ids = ids; state.revision = value.revision;
        return state;
    }
    function nativeMode() { return /\.hta$/i.test(String(window.location.pathname)) && typeof ActiveXObject !== "undefined"; }
    function rootFolder(fso) {
        var path = String(window.location.pathname), host = String(window.location.hostname || "");
        try { path = decodeURIComponent(path); } catch (rawPath) { /* MSHTMLが復号済みのパスを返す場合も使う。 */ }
        path = path.replace(/\//g, "\\");
        if (host && host.toLowerCase() !== "localhost" && !/^\\\\/.test(path)) { path = "\\\\" + host + "\\" + path.replace(/^\\+/, ""); }
        path = path.replace(/^\\([a-z]:\\)/i, "$1");
        if (!/^(?:[a-z]:\\|\\\\)/i.test(path)) { throw new Error("HTAの保存フォルダを取得できません。"); }
        return fso.GetParentFolderName(path);
    }
    function join(folder, name) { return folder.replace(/\\+$/g, "") + "\\" + name; }
    function readLegacy(fso, path) {
        var file = null;
        try { file = fso.OpenTextFile(path, 1, false, -1); return file.ReadAll(); }
        finally { if (file) { try { file.Close(); } catch (ignore) {} } file = null; }
    }
    function read(fso, path) {
        var stream = null;
        try {
            stream = new ActiveXObject("ADODB.Stream"); stream.Type = 2; stream.Charset = "utf-8"; stream.Open();
            stream.LoadFromFile(path); return stream.ReadText(-1);
        } finally { if (stream) { try { stream.Close(); } catch (ignore) {} } stream = null; }
    }
    function create(fso, path, text) {
        var stream = null, output = null;
        try {
            stream = new ActiveXObject("ADODB.Stream"); stream.Type = 2; stream.Charset = "utf-8"; stream.Open();
            stream.WriteText(text.replace(/\n/g, "\r\n")); stream.Position = 0; stream.Type = 1; stream.Position = 3;
            output = new ActiveXObject("ADODB.Stream"); output.Type = 1; output.Open(); output.Write(stream.Read(-1));
            output.SaveToFile(path, 1); // adSaveCreateNotExist / UTF-8のBOMを除いた標準JSON
        } finally {
            if (output) { try { output.Close(); } catch (outputClose) {} }
            if (stream) { try { stream.Close(); } catch (streamClose) {} } output = null; stream = null;
        }
    }
    function signature(fso, path) { var file = fso.GetFile(path); return String(file.DateLastModified) + "|" + String(file.Size); }
    function legacyManaged(text, day) { return String(text).replace(/^\uFEFF/, "").replace(/\r\n/g, "\n").indexOf(legacyHeader(day).split("\n")[0]) === 0; }
    function jsonManaged(text, day) {
        var value;
        try { value = JSON.parse(String(text).replace(/^\uFEFF/, "")); }
        catch (unrelated) { return false; }
        return value && value.schemaVersion === 2 && value.date === day;
    }
    function archive(fso, root, source, name) {
        var old = join(root, "OLD"), match = /^(\d{8})\.(js|json)$/i.exec(name), destination = join(old, name), suffix = 1;
        if (!fso.FolderExists(old)) { fso.CreateFolder(old); }
        while (fso.FileExists(destination)) { destination = join(old, match[1] + "_" + suffix++ + "." + match[2]); }
        fso.MoveFile(source, destination);
    }
    function ensureDay(fso, root, day) {
        var path = join(root, day + ".json"), legacy = join(root, day + ".js"), files, item, match, source, text, initial = null, stage, entries = [], i;
        if (knownDay !== day || !fso.FileExists(path)) {
            files = new Enumerator(fso.GetFolder(root).Files);
            /* 列挙中の移動で取りこぼさないよう、名前・パスの一覧を先に確定する。 */
            for (; !files.atEnd(); files.moveNext()) {
                item = files.item(); entries.push({ Name: String(item.Name), Path: String(item.Path) });
            }
            files = null; item = null;
            for (i = 0; i < entries.length; i++) {
                item = entries[i]; match = /^(\d{8})\.(js|json)$/i.exec(String(item.Name));
                if (!match || match[1] >= day) { continue; }
                source = String(item.Path);
                /* 前日以前のJSONはファイル名だけで判定し、旧JSの形式確認は従来どおり行う。 */
                if (match[2].toLowerCase() === "js" && !legacyManaged(readLegacy(fso, source), match[1])) { continue; }
                archive(fso, root, source, String(item.Name));
            }
            item = null;
        }
        if (fso.FileExists(legacy)) {
            text = readLegacy(fso, legacy);
            if (legacyManaged(text, day)) {
                if (fso.FileExists(path)) { throw new Error("本日の共有JSとJSONが両方あります。全員の旧版フォームを閉じ、READMEの版切替を確認してください。履歴は上書きしません。"); }
                initial = parseLegacy(text, day);
            }
        }
        if (!fso.FileExists(path)) {
            stage = join(join(root, lockName), "new.json");
            create(fso, stage, encode(initial || empty(day))); parse(read(fso, stage), day);
            fso.MoveFile(stage, path);
            if (initial) {
                try { archive(fso, root, legacy, day + ".js"); }
                catch (archiveError) { fso.DeleteFile(path, true); throw archiveError; }
            }
        }
        return path;
    }
    function makeRecord(state, action, id, now) {
        var record = { id: id, supports: [], commits: [] }, key = keyOf(action.organization, action.caName), list = active(state, key), i, row, name = trim(action.supportName), stamp = timeText(now);
        validateIdentity(action);
        if (own(state.ids, "@" + id)) { return record; }
        if (action.kind === "commit") {
            if (committed(state, key)) { return record; }
            if (now.getHours() < 17) { throw new Error("コミットは17:00以降に行えます。"); }
            for (i = 0; i < list.length; i++) {
                row = list[i]; record.supports.push({ id: row.id, organization: row.organization, caName: row.caName, supportName: row.supportName, startedAt: row.startedAt, completedAt: stamp });
            }
            record.commits.push({ organization: trim(action.organization), caName: trim(action.caName), committedAt: stamp });
        } else if (action.kind === "cancelCommit") {
            row = committed(state, key);
            if (!row) { return record; }
            /* 古い画面から、他のフォームが再コミットした情報を取り消さない。 */
            if (row.committedAt !== action.committedAt || state.revision !== action.revision) {
                throw new Error("共有状況が更新されています。最新のコミット状態を確認してから、もう一度取り消してください。");
            }
            record.cancelCommits = [{ organization: row.organization, caName: row.caName, committedAt: row.committedAt }];
        } else if (action.kind === "add") {
            if (committed(state, key)) { throw new Error("このCAはコミット済みです。サポートは追加できません。"); }
            if (!name || name.length > 30) { throw new Error("サポート氏名は1〜30文字で入力してください。"); }
            if (list.length >= 3) { throw new Error("サポートはCAごとに最大3名です。共有状況を確認してください。"); }
            for (i = 0; i < list.length; i++) { if (normalized(list[i].supportName) === normalized(name)) { throw new Error("このCAには、同じ氏名のサポートが登録されています。"); } }
            record.supports.push({ id: id, organization: trim(action.organization), caName: trim(action.caName), supportName: name, startedAt: stamp, completedAt: "" });
        } else if (action.kind === "finish") {
            for (i = 0; i < list.length; i++) {
                row = list[i];
                if (row.id === action.entryId) { record.supports.push({ id: row.id, organization: row.organization, caName: row.caName, supportName: row.supportName, startedAt: row.startedAt, completedAt: stamp }); break; }
            }
        } else { throw new Error("共有操作の種類が不正です。"); }
        return record;
    }
    function persist(fso, path, lock, state, record) {
        var backup = join(lock, "before.json"), stage = join(lock, "after.json"), after, recovery, failure, recoveryError, publishing = false;
        if (!record.supports.length && !record.commits.length && !(record.cancelCommits && record.cancelCommits.length)) { return state; }
        after = parse(encode(state), state.date); apply(after, record);
        fso.CopyFile(path, backup, false);
        try {
            create(fso, stage, encode(after)); parse(read(fso, stage), state.date);
            publishing = true; fso.CopyFile(stage, path, true);
            after = parse(read(fso, path), state.date);
            if (!own(after.ids, "@" + record.id)) { throw new Error("共有JSONへの保存を確認できません。"); }
            return after;
        } catch (writeError) {
            failure = writeError;
            /* 保存済みで応答だけが失敗した場合、同じ操作を重ねて保存しない。 */
            try { after = parse(read(fso, path), state.date); if (own(after.ids, "@" + record.id)) { return after; } }
            catch (readError) {}
            if (!publishing) { throw failure; }
            try { fso.CopyFile(backup, path, true); parse(read(fso, path), state.date); }
            catch (restoreError) {
                recovery = path.replace(/\.json$/i, "_recovery_" + record.id + ".json");
                try { fso.MoveFile(backup, recovery); }
                catch (moveError) { failure.keepLock = true; recovery = backup; }
                recoveryError = new Error("共有JSONを復元できません。更新を止め、管理者に確認してください。\n復元用ファイル：" + recovery + "\n" + (restoreError.message || restoreError.description || String(restoreError)));
                recoveryError.keepLock = !!failure.keepLock; throw recoveryError;
            }
            throw failure;
        } finally {
            /* 控えの保全が必要な場合、後片付けの失敗でkeepLockを失わないよう触らない。 */
            if (!(failure && failure.keepLock)) {
                if (fso.FileExists(stage)) { fso.DeleteFile(stage, true); }
                if (fso.FileExists(backup)) { fso.DeleteFile(backup, true); }
            }
        }
    }
    function notify(error, state) { if (listener) { listener(error, state); } }
    function finish(error, state) {
        var job = current;
        current = null; retryTimer = null;
        if (state) { lastState = state; }
        if (job && job.callback) { job.callback(error, state); }
        notify(error, state); pump();
    }
    function attempt() {
        var fso = null, root, lock, ownsLock = false, day = dayKey(new Date()), path, state = null, record, error = null, sig, keep = false;
        if (!current || stopped) { return; }
        if (!nativeMode()) {
            try {
                state = lastState && lastState.date === day ? lastState : empty(day);
                if (current.action) { record = makeRecord(state, current.action, current.id, new Date()); if (record.supports.length || record.commits.length || (record.cancelCommits && record.cancelCommits.length)) { apply(state, record); } }
            } catch (previewError) { error = previewError; }
            finish(error, state); return;
        }
        try {
            fso = new ActiveXObject("Scripting.FileSystemObject"); root = rootFolder(fso); lock = join(root, lockName);
            try { fso.CreateFolder(lock); ownsLock = true; }
            catch (lockError) {
                /* 相手が直後に解除した場合、FolderExistsはfalseでも作成競合を再試行する。 */
                if (!fso.FolderExists(lock) && (Number(lockError.number) & 65535) !== 58) { throw lockError; }
                if (current.attempts++ >= retryLimit) { throw new Error("共有ファイルが他のフォームで更新中です。少し待って再度操作してください。\n続く場合はREADMEの共有ファイルの復旧を確認してください。"); }
                retryTimer = window.setTimeout(attempt, 150 + Math.floor(Math.random() * 100)); return;
            }
            path = ensureDay(fso, root, day); sig = signature(fso, path);
            if (!current.action && lastState && knownDay === day && sig === lastSignature) { state = lastState; }
            else { state = parse(read(fso, path), day); }
            if (current.action) {
                if (current.date !== day) { throw new Error("日付が変わりました。当日の共有状況を確認してからもう一度操作してください。"); }
                record = makeRecord(state, current.action, current.id, new Date());
                state = persist(fso, path, lock, state, record);
            }
            knownDay = day; lastSignature = signature(fso, path);
        } catch (operationError) { error = operationError; keep = !!operationError.keepLock; }
        finally {
            if (ownsLock && !keep) { try { fso.DeleteFolder(lock, true); } catch (unlockError) { if (!error) { error = new Error("共有履歴は保存されましたが、更新ロックを解除できませんでした。\n" + lock); } } }
            fso = null;
        }
        finish(error, state);
    }
    function pump() { if (!stopped && !current && queue.length) { current = queue.shift(); attempt(); } }
    function request(action, callback) {
        if (stopped) { if (callback) { callback(new Error("共有ファイルの準備ができていません。フォームを開き直してください。"), lastState); } return; }
        queue.push({ action: action, callback: callback, id: session + "-" + (++sequence).toString(36), date: dayKey(new Date()), attempts: 0 }); pump();
    }
    function stop() {
        stopped = true;
        if (pollTimer) { window.clearInterval(pollTimer); pollTimer = null; }
        if (retryTimer) { window.clearTimeout(retryTimer); retryTimer = null; }
        queue = []; current = null; listener = null;
    }
    function start(callback) {
        stop(); stopped = false; listener = callback; knownDay = ""; lastSignature = ""; lastState = null;
        request(null, null);
        pollTimer = window.setInterval(function () { if (!current && !queue.length) { request(null, null); } }, 3000);
    }
    return {
        start: start, stop: stop, refresh: function () { request(null, null); },
        addSupport: function (ca, name, callback) { request({ kind: "add", organization: ca.organization, caName: ca.name, supportName: name }, callback); },
        finishSupport: function (ca, entryId, callback) { request({ kind: "finish", organization: ca.organization, caName: ca.name, entryId: entryId }, callback); },
        commitCA: function (ca, callback) { request({ kind: "commit", organization: ca.organization, caName: ca.name }, callback); },
        cancelCommitCA: function (ca, revision, callback) { request({ kind: "cancelCommit", organization: ca.organization, caName: ca.name, committedAt: ca.committedAt, revision: revision }, callback); },
        keyOf: keyOf, dayKey: dayKey, activeForCA: active, committedForCA: committed, parseDay: parse, headerForDay: function (day) { return encode(empty(day)); }, parseLegacyDay: parseLegacy
    };
}());

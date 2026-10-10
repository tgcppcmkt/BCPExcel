/* Browser edition: selected files only; no external libraries or remote requests. */
var WorkloadBrowserStore = (function () {
    var opening = null;
    function open() {
        if (!opening) {
            opening = new Promise(function (resolve, reject) {
                if (!window.indexedDB) { reject(new Error("ブラウザに保存先を記憶できません。保存設定を確認してください。")); return; }
                var request = window.indexedDB.open("WorkloadRankingHTML", 1);
                request.onupgradeneeded = function () { request.result.createObjectStore("handles"); };
                request.onsuccess = function () { resolve(request.result); };
                request.onerror = function () { reject(request.error); };
                request.onblocked = function () { reject(new Error("保存設定の更新が止まっています。他の業務量モニターを閉じて再度開いてください。")); };
            });
        }
        return opening;
    }
    async function access(key, value, writing) {
        var db = await open();
        return new Promise(function (resolve, reject) {
            var tx = db.transaction("handles", writing ? "readwrite" : "readonly"), store = tx.objectStore("handles");
            var request = writing ? store.put(value, key) : store.get(key);
            tx.oncomplete = function () { resolve(writing ? value : request.result); };
            tx.onerror = tx.onabort = function () { reject(tx.error || request.error || new Error("ブラウザの保存設定を記憶できません。")); };
        });
    }
    return { get: function (key) { return access(key); }, set: function (key, value) { return access(key, value, true); } };
}());

var WorkloadWorkbook = (function () {
    var MAX_FILE = 64 * 1024 * 1024, MAX_XML = 32 * 1024 * 1024;
    var crcTable = Array.from({ length: 256 }, function (_, value) {
        for (var bit = 0; bit < 8; bit++) { value = (value >>> 1) ^ ((value & 1) ? 0xedb88320 : 0); }
        return value >>> 0;
    });
    function error(message) { return new Error("振分表：" + message); }
    function crc32(bytes) {
        var crc = 0xffffffff;
        for (var i = 0; i < bytes.length; i++) {
            crc = (crc >>> 8) ^ crcTable[(crc ^ bytes[i]) & 255];
        }
        return (crc ^ 0xffffffff) >>> 0;
    }
    function packagePath(base, target) {
        if (/[\\\x00-\x1f]/.test(target) || /^[a-z]+:/i.test(target)) { throw error("ブック内の参照先が不正です。"); }
        var parts = (target.charAt(0) === "/" ? target.slice(1) : base + target).split("/"), result = [];
        for (var item of parts) {
            if (!item || item === ".") { continue; }
            if (item === "..") { if (!result.length) { throw error("ブック内の参照が範囲外です。"); } result.pop(); }
            else { result.push(item); }
        }
        return result.join("/");
    }
    function nodes(parent, tag) { return Array.from(parent.getElementsByTagNameNS("*", tag)); }
    function child(parent, tag) { return nodes(parent, tag)[0] || null; }
    function xml(bytes) {
        var encoding = bytes[0] === 255 && bytes[1] === 254 ? "utf-16le" : bytes[0] === 254 && bytes[1] === 255 ? "utf-16be" : "utf-8";
        var text = new TextDecoder(encoding, { fatal: true }).decode(bytes);
        if (/<!DOCTYPE|<!ENTITY/i.test(text)) { throw error("外部定義を含むXMLは読み込みません。"); }
        var doc = new DOMParser().parseFromString(text, "application/xml");
        if (nodes(doc, "parsererror").length) { throw error("Excel内部のXMLが壊れています。"); }
        return doc;
    }
    function textRuns(node) {
        return nodes(node, "t").filter(function (t) {
            var p = t.parentNode;
            while (p && p !== node) { if (p.localName === "rPh") { return false; } p = p.parentNode; }
            return true;
        }).map(function (t) { return t.textContent; }).join("");
    }
    async function zip(file) {
        if (file.size > MAX_FILE) { throw error("64MBを超えています。取込用のコピーを小さくして選択してください。"); }
        var bytes = new Uint8Array(await file.arrayBuffer()), view = new DataView(bytes.buffer), end = -1;
        if (bytes.length < 22) { throw error(".xlsx または .xlsm 形式ではありません。"); }
        for (var i = bytes.length - 22; i >= Math.max(0, bytes.length - 65557); i--) {
            if (view.getUint32(i, true) === 0x06054b50 && i + 22 + view.getUint16(i + 20, true) === bytes.length) { end = i; break; }
        }
        if (end < 0) { throw error(".xlsx／.xlsm を選択してください。.xls・.xlsb・パスワード付きブックは対応していません。"); }
        var count = view.getUint16(end + 10, true), offset = view.getUint32(end + 16, true), directorySize = view.getUint32(end + 12, true);
        if (view.getUint16(end + 4, true) || view.getUint16(end + 6, true) || count === 65535 || count > 10000 || offset + directorySize > end) { throw error("分割・ZIP64形式または大きすぎるブックです。"); }
        var entries = new Map(), total = 0;
        for (i = 0; i < count; i++) {
            if (offset + 46 > end || view.getUint32(offset, true) !== 0x02014b50) { throw error("Excel内部のファイル一覧が壊れています。"); }
            var length = view.getUint16(offset + 28, true), extra = view.getUint16(offset + 30, true), comment = view.getUint16(offset + 32, true);
            if (offset + 46 + length + extra + comment > end) { throw error("Excel内部のファイル名が壊れています。"); }
            var name = new TextDecoder("utf-8", { fatal: true }).decode(bytes.subarray(offset + 46, offset + 46 + length));
            if (name.charAt(0) === "/" || name.split("/").indexOf("..") >= 0 || /[\\\x00-\x1f]/.test(name) || entries.has(name)) { throw error("Excel内部のファイル名が不正です。"); }
            var entry = { flags: view.getUint16(offset + 8, true), method: view.getUint16(offset + 10, true),
                crc: view.getUint32(offset + 16, true), size: view.getUint32(offset + 24, true),
                compressed: view.getUint32(offset + 20, true), offset: view.getUint32(offset + 42, true) };
            total += entry.size; if (total > 256 * 1024 * 1024) { throw error("展開後のブックが大きすぎます。"); }
            entries.set(name, entry); offset += 46 + length + extra + comment;
        }
        async function read(name, optional) {
            var e = entries.get(name); if (!e) { if (optional) { return null; } throw error("必要な内部ファイルがありません：" + name); }
            if (e.flags & 1 || e.size > MAX_XML || e.offset + 30 > bytes.length || view.getUint32(e.offset, true) !== 0x04034b50) { throw error("暗号化または壊れた内部ファイルです：" + name); }
            var start = e.offset + 30 + view.getUint16(e.offset + 26, true) + view.getUint16(e.offset + 28, true);
            if (start + e.compressed > bytes.length) { throw error("内部ファイルの長さが不正です。"); }
            var compressed = bytes.subarray(start, start + e.compressed), result;
            if (e.method === 0) { result = compressed; }
            else if (e.method === 8) {
                if (typeof DecompressionStream === "undefined") { throw error("このブラウザではExcelを展開できません。通常のEdgeの最新版で開いてください。"); }
                var stream;
                try { stream = new Blob([compressed]).stream().pipeThrough(new DecompressionStream("deflate-raw")); }
                catch (unsupported) { throw error("このEdgeではExcelの展開形式に対応していません。ブラウザの更新を確認してください。"); }
                var reader = stream.getReader(), chunks = [], size = 0;
                try {
                    while (true) {
                        var part = await reader.read(); if (part.done) { break; }
                        size += part.value.length;
                        if (size > MAX_XML || size > e.size) { throw error("内部ファイルの展開サイズが不正です。"); }
                        chunks.push(part.value);
                    }
                } catch (badStream) { await reader.cancel().catch(function () {}); throw badStream; }
                result = new Uint8Array(size); var cursor = 0;
                for (var chunk of chunks) { result.set(chunk, cursor); cursor += chunk.length; }
            } else { throw error("対応していないExcelの圧縮形式です。"); }
            if (result.length !== e.size || crc32(result) !== e.crc) { throw error("内部ファイルの破損を検出しました：" + name); }
            return result;
        }
        return { read: read };
    }
    async function read(file, config) {
        if (!/\.(xlsx|xlsm)$/i.test(file.name)) { throw error(".xlsx または .xlsm を選択してください。"); }
        var archive = await zip(file), workbookPath = "xl/workbook.xml", rootRels = await archive.read("_rels/.rels", true);
        if (rootRels) {
            for (var rel of nodes(xml(rootRels), "Relationship")) {
                if (/\/officeDocument$/.test(rel.getAttribute("Type")) && rel.getAttribute("TargetMode") !== "External") { workbookPath = packagePath("", rel.getAttribute("Target")); }
            }
        }
        var base = workbookPath.slice(0, workbookPath.lastIndexOf("/") + 1), book = xml(await archive.read(workbookPath));
        var relPath = base + "_rels/" + workbookPath.slice(base.length) + ".rels", rels = new Map();
        for (rel of nodes(xml(await archive.read(relPath)), "Relationship")) { rels.set(rel.getAttribute("Id"), rel); }
        var targetSheet = nodes(book, "sheet").find(function (s) { return s.getAttribute("name") === config.assignment.sheetName; });
        if (!targetSheet) { throw error("「" + config.assignment.sheetName + "」シートが見つかりません。"); }
        var relationId = targetSheet.getAttribute("r:id") || targetSheet.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
        var sheetRel = rels.get(relationId);
        if (!sheetRel || sheetRel.getAttribute("TargetMode") === "External") { throw error("対象シートの参照が不正です。"); }
        var sheetPath = packagePath(base, sheetRel.getAttribute("Target")), strings = [], shared = null;
        for (rel of rels.values()) { if (/\/sharedStrings$/.test(rel.getAttribute("Type")) && rel.getAttribute("TargetMode") !== "External") { shared = await archive.read(packagePath(base, rel.getAttribute("Target"))); } }
        if (shared) { strings = nodes(xml(shared), "si").map(textRuns); }
        var sheet = xml(await archive.read(sheetPath)), cols = config.assignment.columns, maximum = 0;
        for (var name of Object.keys(cols)) { maximum = Math.max(maximum, WorkloadImport.columnIndex(cols[name])); }
        var sparse = new Map(), formulas = new Map(), lastRow = 0;
        for (var cell of nodes(sheet, "c")) {
            var reference = cell.getAttribute("r"), match = /^([A-Z]{1,3})([1-9]\d*)$/.exec(reference);
            if (!match) { throw error("セル位置が不正です。"); }
            var column = WorkloadImport.columnIndex(match[1]), row = Number(match[2]);
            if (row > 1048576) { throw error("セルの行番号が範囲外です。"); }
            if (column > maximum) { continue; }
            var v = child(cell, "v"), f = child(cell, "f"), inline = child(cell, "is"), type = cell.getAttribute("t"), value = v ? v.textContent : "";
            if (type === "s") {
                var index = Number(value); if (!/^\d+$/.test(value) || !Number.isSafeInteger(index) || index >= strings.length) { throw error(reference + " の文字列参照が不正です。"); } value = strings[index];
            } else if (type === "inlineStr") { value = inline ? textRuns(inline) : ""; }
            else if (type === "e" && Object.values(cols).indexOf(match[1]) >= 0) { throw error(reference + " にExcelのエラーがあります：" + value); }
            else if (type === "b") { value = value === "1" ? "true" : "false"; }
            if (f) {
                formulas.set(reference, f.textContent);
                if (!v && !inline && column !== WorkloadImport.columnIndex(cols.calib) && column !== WorkloadImport.columnIndex(cols.canvas) && Object.values(cols).indexOf(match[1]) >= 0) {
                    throw error(reference + " の数式結果が保存されていません。Excelで開いて保存したブックを選択してください。");
                }
            }
            if (!value && !f) { continue; }
            if (!sparse.has(row)) { sparse.set(row, Array(maximum + 1).fill("")); }
            sparse.get(row)[column] = value;
            if ((column === WorkloadImport.columnIndex(cols.organization) || column === WorkloadImport.columnIndex(cols.caName)) && String(value).trim()) { lastRow = Math.max(lastRow, row); }
        }
        if (lastRow < config.assignment.startRow) { throw error("「" + config.assignment.sheetName + "」に取込対象のCAがありません。"); }
        if (lastRow > 200000) { throw error("対象シートの最終行が200,000行を超えています。"); }
        var rows = Array.from({ length: lastRow }, function (_, i) { return sparse.get(i + 1) || Array(maximum + 1).fill(""); }), links = {};
        var sheetBase = sheetPath.slice(0, sheetPath.lastIndexOf("/") + 1), linkRelations = new Map();
        var linkBytes = await archive.read(sheetBase + "_rels/" + sheetPath.slice(sheetBase.length) + ".rels", true);
        if (linkBytes) { for (rel of nodes(xml(linkBytes), "Relationship")) { if (/\/hyperlink$/.test(rel.getAttribute("Type"))) { linkRelations.set(rel.getAttribute("Id"), rel.getAttribute("Target")); } } }
        function linkAt(ref, url) {
            var m = /^([A-Z]{1,3})([1-9]\d*)$/.exec(ref); if (!m || Number(m[2]) > lastRow) { return; }
            var property = m[1] === cols.calib ? "calib" : m[1] === cols.canvas ? "canvas" : "";
            if (property && WorkloadImport.validURL(url)) { var key = "@" + m[2]; if (!links[key]) { links[key] = {}; } if (!links[key][property]) { links[key][property] = url; } }
        }
        for (var hyperlink of nodes(sheet, "hyperlink")) {
            var id = hyperlink.getAttribute("r:id") || hyperlink.getAttributeNS("http://schemas.openxmlformats.org/officeDocument/2006/relationships", "id");
            var url = linkRelations.get(id); if (url) { linkAt(hyperlink.getAttribute("ref").split(":")[0], url); }
        }
        for (var pair of formulas) {
            var formula = /^=?\s*HYPERLINK\s*\(\s*"((?:[^"]|"")*)"\s*[,;)]/i.exec(pair[1]);
            if (formula) { linkAt(pair[0], formula[1].replace(/""/g, '"')); }
        }
        var warnings = [];
        for (var rowIndex = config.assignment.startRow - 1; rowIndex < lastRow; rowIndex++) {
            for (var property of ["calib", "canvas"]) {
                var columnName = cols[property], cellRef = columnName + (rowIndex + 1);
                var savedValue = String(rows[rowIndex][WorkloadImport.columnIndex(columnName)] || "").trim();
                if (!(links["@" + (rowIndex + 1)] || {})[property] && !WorkloadImport.validURL(savedValue) && (savedValue || formulas.has(cellRef))) {
                    warnings.push(cellRef + "：リンク先URLを取得できません。URLを直接記入するか、通常のハイパーリンクを設定してください。");
                }
            }
        }
        return { rows: rows, linksByRow: links, warnings: warnings };
    }
    return { read: read };
}());

var WorkloadBrowser = (function () {
    var selected = {}, roots = {}, starts = {}, remembered = null, choosing = false, shared = null;
    function node(id) { return document.getElementById(id); }
    function info(message, bad) { var el = node("browserStatus"); if (el) { el.textContent = message; el.className = "browser-note" + (bad ? " error" : ""); } }
    function fail(error) { if (error.name !== "AbortError") { var message = error.message || String(error); info(message, true); window.alert(message); } }
    function allowed() { return !choosing && (!window.WorkloadUI || WorkloadUI.canChooseBrowserFiles()); }
    function display(kind) {
        var el = node(kind + "FileName"), choice = selected[kind];
        if (el) { el.textContent = choice ? choice.name : "未選択"; el.title = el.textContent; }
    }
    async function rememberChoice(kind, handle) {
        var file = await handle.getFile();
        if (!(kind === "assignment" ? /\.(xlsx|xlsm)$/i : /\.csv$/i).test(file.name)) { throw new Error(kind === "assignment" ? "振分表は .xlsx または .xlsm を選択してください。" : "CSVファイルを選択してください。"); }
        selected[kind] = { handle: handle, name: file.name }; display(kind);
        await WorkloadBrowserStore.set("file-" + kind, handle);
    }
    async function dateStart(kind) {
        if (!roots[kind]) { return; }
        var date = new Date(), month = date.getFullYear() + ("0" + (date.getMonth() + 1)).slice(-2), day = ("0" + (date.getMonth() + 1)).slice(-2) + ("0" + date.getDate()).slice(-2);
        var folder = roots[kind], parts = kind === "weekend" ? [String(date.getFullYear()).slice(-2) + day] : kind === "unread" ? [month, day] : [month];
        try { for (var part of parts) { folder = await folder.getDirectoryHandle(part); } starts[kind] = { handle: folder, parts: parts, exact: true, day: date.toDateString() }; }
        catch (error) { if (error.name !== "NotFoundError") { throw error; } starts[kind] = { handle: roots[kind], parts: [], exact: false, day: date.toDateString() }; }
    }
    function pickerOptions(kind) {
        var opts = { multiple: false, id: "workload-" + kind, types: kind === "assignment" ?
            [{ description: "振分表Excel", accept: { "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": [".xlsx"], "application/vnd.ms-excel.sheet.macroEnabled.12": [".xlsm"] } }] :
            [{ description: "CSV", accept: { "text/csv": [".csv"] } }] };
        var start = starts[kind];
        if (start && start.day === new Date().toDateString()) { opts.startIn = start.handle; }
        else if (roots[kind]) { opts.startIn = roots[kind]; }
        else if (selected[kind] && selected[kind].handle) { opts.startIn = selected[kind].handle; }
        return opts;
    }
    async function acceptChoice(kind, handle) {
        if (roots[kind]) {
            await dateStart(kind);
            var start = starts[kind], relative = await roots[kind].resolve(handle);
            if (!relative || start && start.exact && start.day === new Date().toDateString() && relative.slice(0, -1).join("/") !== start.parts.join("/")) { throw new Error((start && !start.exact ? "設定した親フォルダ内" : "設定した当日／当月フォルダ内") + "のCSVを選択してください。"); }
        }
        await rememberChoice(kind, handle);
    }
    function needsLogGesture() {
        return !!(window.navigator && window.navigator.userActivation && !window.navigator.userActivation.isActive);
    }
    function waitForLogGesture(message, caption, report) {
        report(message);
        var dialog = node("logSelectionDialog"), yes = node("logSelectionContinue"), no = node("logSelectionCancel");
        node("logSelectionMessage").textContent = message; yes.textContent = caption;
        return new Promise(function (resolve, reject) {
            function finish(accepted) {
                yes.onclick = no.onclick = dialog.oncancel = null; dialog.close();
                if (accepted) { resolve(); }
                else { reject(new DOMException("ログ集計をキャンセルしました。", "AbortError")); }
            }
            yes.onclick = function () { finish(true); }; no.onclick = function () { finish(false); };
            dialog.oncancel = function (event) { event.preventDefault(); finish(false); };
            dialog.showModal();
        });
    }
    async function logGestureAction(action, message, caption, report) {
        if (needsLogGesture()) { await waitForLogGesture(message, caption, report); }
        try { return await action(); }
        catch (error) {
            if (error.name !== "SecurityError") { throw error; }
            /* 操作許可が直前に切れた場合も、利用者のクリックから1度だけ再開する。 */
            await waitForLogGesture(message, caption, report);
            return await action();
        }
    }
    async function logPermission(handle, mode, label, report) {
        if (await handle.queryPermission({ mode: mode }) === "granted") { return; }
        var access = await logGestureAction(function () { return handle.requestPermission({ mode: mode }); },
            "「" + label + "」のアクセス許可を確認します。「許可を確認」を押してください。", "許可を確認", report);
        if (access !== "granted") { throw new Error("「" + label + "」への" + (mode === "readwrite" ? "読み書き" : "読込") + "を許可してください。"); }
    }
    async function reloadForLogs(report) {
        report("共有情報を更新しています。");
        await logPermission(remembered.handle, "readwrite", "共有", report);
        while (WorkloadShare.isBusy()) { await new Promise(function (resolve) { window.setTimeout(resolve, 50); }); }
        if (!WorkloadShare.isConnected()) { await WorkloadShare.connect(remembered.handle, remembered.key); }
        else { await WorkloadShare.refresh(); }
        /* refreshの読込・保存失敗は状態通知で返るため、CSV選択前に確認する。 */
        if (shared && shared.error) { throw new Error("共有情報を更新できません。\n" + shared.error); }
    }
    async function pickLogFile(kind, report) {
        var options = pickerOptions(kind);
        var handles = await logGestureAction(function () { return window.showOpenFilePicker(options); },
            "CSV選択画面を開きます。「CSVを選択」を押してください。", "CSVを選択", report);
        return handles[0];
    }
    async function selectLogs(progress) {
        if (choosing) { throw new Error("ファイル選択が完了してから集計してください。"); }
        /* File System Accessがない環境では、従来どおり選択済みファイルを集計する。 */
        if (typeof window.showOpenFilePicker !== "function") { return { cancelled: false }; }
        var missing = [];
        if (!remembered) { missing.push("共有"); }
        if (!roots.unread) { missing.push("未読"); }
        if (!roots.flag) { missing.push("フラグ"); }
        if (!selected.assignment) { missing.push("振分表"); }
        if (missing.length) { throw new Error("初期設定が完了していません。画面上部の「設定」で「" + missing.join("」「") + "」を設定してください。"); }
        var report = progress || function () {};
        choosing = true;
        try {
            await reloadForLogs(report);
            for (var kind of ["unread", "flag"]) {
                if (!roots[kind]) { continue; }
                report((kind === "unread" ? "未読" : "フラグ") + "の読込フォルダを確認しています。");
                await logPermission(roots[kind], "read", kind === "unread" ? "未読" : "フラグ", report);
                await dateStart(kind);
            }
            report("1/2 未読CSVを選択してください。");
            var unread = await pickLogFile("unread", report);
            /* 次の選択画面は、未読の読込・設定保存で待たせずに開く。 */
            report("2/2 フラグCSVを選択してください。");
            var flag = await pickLogFile("flag", report);
            var assignment = selected.assignment;
            if (assignment.handle) { await logPermission(assignment.handle, "read", "振分表", report); }
            await acceptChoice("unread", unread);
            await acceptChoice("flag", flag);
            return { cancelled: false };
        } catch (error) {
            if (error.name === "AbortError") { return { cancelled: true }; }
            throw error;
        } finally { choosing = false; }
    }
    async function selectWeekendLogs(progress) {
        if (choosing) { throw new Error("ファイル選択が完了してから集計してください。"); }
        if (typeof window.showOpenFilePicker !== "function") { throw new Error("土日ログの選択には、通常のEdgeでHTMLを開いてください。"); }
        var missing = [], report = progress || function () {}, handles = [];
        if (!remembered) { missing.push("共有"); }
        if (!roots.weekend) { missing.push("土日ログ"); }
        if (!selected.assignment) { missing.push("振分表"); }
        if (missing.length) { throw new Error("初期設定が完了していません。画面上部の「設定」で「" + missing.join("」「") + "」を設定してください。"); }
        choosing = true;
        try {
            await reloadForLogs(report);
            await logPermission(roots.weekend, "read", "土日ログ", report);
            await dateStart("weekend");
            var selectedDay = new Date().toDateString(), options = pickerOptions("weekend");
            for (var i = 0; i < 3; i++) {
                report("土日ログ " + (i + 1) + "/3：CSVを選択してください。");
                var choice = await logGestureAction(function () { return window.showOpenFilePicker(options); },
                    "土日ログ " + (i + 1) + "/3 の選択画面を開きます。「CSVを選択」を押してください。", "CSVを選択", report);
                handles.push(choice[0]);
            }
            for (i = 0; i < handles.length; i++) {
                var relative = await roots.weekend.resolve(handles[i]);
                if (!relative || (starts.weekend.exact && relative.slice(0, -1).join("/") !== starts.weekend.parts.join("/"))) { throw new Error((starts.weekend.exact ? "土日ログの当日フォルダ内" : "「土日ログ」に設定した親フォルダ内") + "のCSVを選択してください。"); }
                if (!/\.csv$/i.test(handles[i].name)) { throw new Error("土日ログにはCSVファイルを選択してください。"); }
                for (var j = 0; j < i; j++) {
                    if (await handles[i].isSameEntry(handles[j])) { throw new Error("同じCSVが選択されています。異なる3つのCSVを選択してください。"); }
                }
            }
            if (selected.assignment.handle) { await logPermission(selected.assignment.handle, "read", "振分表", report); }
            if (selectedDay !== new Date().toDateString()) { throw new Error("選択中に日付が変わりました。もう一度「ログ集計」で当日のCSVを選択してください。"); }
            return { cancelled: false, handles: handles, day: selectedDay };
        } catch (error) {
            if (error.name === "AbortError") { return { cancelled: true }; }
            throw error;
        } finally { choosing = false; }
    }
    async function runWeekend(previousStaff, selection, progress) {
        if (typeof WorkloadWeekend === "undefined") { throw new Error("workload-weekend.jsをHTMLと同じフォルダに配置してください。"); }
        if (!selection || !selection.handles || selection.handles.length !== 3) { throw new Error("土日ログのCSVを3つ選択してください。"); }
        var config = WorkloadImport.validateConfig(), report = progress || function () {}, logs = [], files = [];
        await dateStart("weekend");
        for (var i = 0; i < 3; i++) {
            var handle = selection.handles[i], relative = await roots.weekend.resolve(handle);
            if (selection.day !== new Date().toDateString() || !relative || (starts.weekend.exact && relative.slice(0, -1).join("/") !== starts.weekend.parts.join("/"))) { throw new Error("当日の土日ログCSVをもう一度選択してください。"); }
            report("土日ログ " + (i + 1) + "/3 を読み込んでいます。");
            var file = await handle.getFile(), decoded = await csv(file, "auto"); files.push(file.name);
            logs.push(WorkloadWeekend.readLog(decoded.rows, file.name));
        }
        report("振分表を読み込んでいます。");
        var assignmentFile = await fileOf("assignment"), roster = await WorkloadWorkbook.read(assignmentFile, config);
        var result = WorkloadWeekend.build(logs, roster, previousStaff, config);
        if (selection.day !== new Date().toDateString()) { throw new Error("読み込み中に日付が変わりました。もう一度集計してください。"); }
        result.weekend.day = selection.day; result.importedAt = new Date();
        result.sources = { weekendFiles: files, assignmentFile: assignmentFile.name, sheetName: config.assignment.sheetName };
        return result;
    }
    async function choose(kind) {
        if (!allowed()) { info("入力・保存・集計が完了してからファイルを選択してください。", true); return; }
        if (typeof window.showOpenFilePicker !== "function") {
            var input = node("browserFileInput"); input.accept = kind === "assignment" ? ".xlsx,.xlsm" : ".csv"; input.value = "";
            input.onchange = function () { var file = input.files[0]; if (file) { selected[kind] = { file: file, name: file.name }; display(kind); } }; input.click(); return;
        }
        choosing = true;
        try {
            var handles = await window.showOpenFilePicker(pickerOptions(kind));
            await acceptChoice(kind, handles[0]); info("ファイルを選択しました。「ログ集計」で未読CSV・フラグCSVを選択してください。");
        } catch (error) { fail(error); } finally { choosing = false; }
    }
    async function chooseRoot(kind) {
        if (!allowed()) { info("入力・保存・集計が完了してから設定してください。", true); return; }
        if (!window.showDirectoryPicker) { fail(new Error("フォルダ選択が利用できません。通常のEdgeで開き、ブラウザ設定を管理担当へ確認してください。")); return; }
        choosing = true;
        try {
            var handle = await window.showDirectoryPicker({ id: "workload-root-" + kind, mode: "read" });
            roots[kind] = handle; await WorkloadBrowserStore.set("root-" + kind, handle);
            node(kind + "RootName").textContent = handle.name; node(kind + "RootName").title = handle.name; await dateStart(kind);
            info("読込フォルダを設定しました。「ログ集計」でCSVを選択してください。");
        } catch (error) { fail(error); } finally { choosing = false; }
    }
    async function connectShared() {
        if (!allowed()) { info("入力・保存・集計が完了してから共有先を接続してください。", true); return; }
        if (!window.showDirectoryPicker || !window.isSecureContext) { fail(new Error("共有フォルダを接続できません。通常のEdgeでHTMLファイルを開き、ブラウザのファイルアクセス設定を確認してください。")); return; }
        choosing = true; node("connectShared").disabled = true;
        try {
            var handle, key;
            handle = await window.showDirectoryPicker({ id: "workload-shared", mode: "readwrite" });
            key = remembered && await handle.isSameEntry(remembered.handle) ? remembered.key : "shared-root-" + window.crypto.randomUUID();
            await WorkloadShare.connect(handle, key);
            remembered = { handle: handle, key: key }; await WorkloadBrowserStore.set("shared-current", remembered);
            node("connectShared").textContent = "共有";
        } catch (error) { fail(error); }
        finally { choosing = false; node("connectShared").disabled = false; }
    }
    async function retryShared() {
        if (choosing || window.WorkloadUI && !WorkloadUI.canRetryBrowserWrites()) { info("保存・集計が完了してから再読み込みしてください。", true); return; }
        try {
            if (!remembered) { await connectShared(); return; }
            if (await remembered.handle.requestPermission({ mode: "readwrite" }) !== "granted") { throw new Error("共有フォルダへの読み書きを許可してください。"); }
            if (!WorkloadShare.isConnected()) { await WorkloadShare.connect(remembered.handle, remembered.key); }
            else { await WorkloadShare.refresh(); }
        } catch (error) { fail(error); }
    }
    async function fileOf(kind) {
        var choice = selected[kind];
        if (!choice) { throw new Error(kind === "assignment" ? "「設定」の「振分表」から振分表を指定してください。" : "「ログ集計」で" + (kind === "unread" ? "未読CSV" : "フラグCSV") + "を選択してください。"); }
        if (choice.handle) {
            if (await choice.handle.queryPermission({ mode: "read" }) !== "granted") { throw new Error(choice.name + " の読込許可が必要です。" + (kind === "assignment" ? "「設定」の「振分表」" : "「ログ集計」") + "からもう一度選択してください。"); }
            if (roots[kind]) {
                await dateStart(kind);
                var start = starts[kind], relative = await roots[kind].resolve(choice.handle);
                if (!relative || (start.exact && relative.slice(0, -1).join("/") !== start.parts.join("/"))) { throw new Error((start.exact ? "当日／当月フォルダ内の " : "設定した親フォルダ内の ") + choice.name + " を選び直してください。"); }
            }
            return await choice.handle.getFile();
        }
        return choice.file;
    }
    async function csv(file, charset) {
        if (file.size > 64 * 1024 * 1024) { throw new Error(file.name + " が64MBを超えています。"); }
        var bytes = new Uint8Array(await file.arrayBuffer()), encoding = charset === "auto" ? WorkloadImport.detectCharset(bytes) : charset;
        var labels = { unicode: "utf-16le", unicodeFFFE: "utf-16be" }, text = new TextDecoder(labels[encoding] || encoding, { fatal: true }).decode(bytes);
        return { rows: WorkloadImport.parseCSV(text), charset: encoding };
    }
    async function runWeekday(previousStaff, progress) {
        var config = WorkloadImport.validateConfig(), report = progress || function () {};
        report("1/3 未読CSVを読み込んでいます。"); var unreadFile = await fileOf("unread"), unread = await csv(unreadFile, config.unread.charset);
        var unreadLog = WorkloadImport.readLogRows(unread.rows, "unread", config, "未読CSV");
        report("2/3 フラグCSVを読み込んでいます。"); var flagFile = await fileOf("flag"), flag = await csv(flagFile, config.flag.charset);
        var flagLog = WorkloadImport.readLogRows(flag.rows, "flag", config, "フラグCSV");
        report("3/3 振分表を読み込んでいます。"); var assignmentFile = await fileOf("assignment"), roster = await WorkloadWorkbook.read(assignmentFile, config);
        report("CSVと振分表を照合しています。");
        var result = WorkloadImport.buildFromLogs(unreadLog, flagLog, roster.rows, previousStaff, config, roster.linksByRow);
        result.warnings = roster.warnings;
        result.sources = { unreadFile: unreadFile.name, flagFile: flagFile.name, assignmentFile: assignmentFile.name,
            sheetName: config.assignment.sheetName, unreadCharset: unread.charset, flagCharset: flag.charset };
        result.importedAt = new Date(); result.hourlyRate = config.capacity.hourlyRate; return result;
    }
    function sharedStatus(value) {
        shared = value;
        if (value.connected) { node("sharedRootName").textContent = value.folder; node("sharedRootName").title = value.folder; }
        info((value.error || (value.connected ? "" : "「設定」の「共有」で接続してください。") +
            (value.pending ? " ／ 保存待ち " + value.pending + "件（完了まで画面を開いてください）" : "") +
            (value.rejected.length ? " ／ 同時操作による未反映 " + value.rejected.length + "件" : "")).replace(/^ ／ /, ""), !!value.error || !!value.rejected.length);
        var box = node("browserConflict");
        if (box) { box.textContent = value.rejected.map(function (x) { return x.stamp + "　" + x.organization + "／" + x.caName + "：" + x.message; }).join("\n"); }
    }
    function download(text, filename) {
        var url = URL.createObjectURL(new Blob([text], { type: "application/json;charset=utf-8" })), a = document.createElement("a");
        a.href = url; a.download = filename; document.body.appendChild(a); a.click(); a.remove(); window.setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
    }
    async function init() {
        try {
            for (var kind of ["unread", "flag", "assignment"]) {
                var handle = await WorkloadBrowserStore.get("file-" + kind);
                if (handle) { selected[kind] = { handle: handle, name: handle.name }; display(kind); }
            }
            for (kind of ["unread", "flag", "weekend"]) {
                handle = await WorkloadBrowserStore.get("root-" + kind);
                if (handle) { roots[kind] = handle; node(kind + "RootName").textContent = handle.name; node(kind + "RootName").title = handle.name; await dateStart(kind).catch(function () {}); }
            }
            remembered = await WorkloadBrowserStore.get("shared-current");
            if (remembered) { node("sharedRootName").textContent = remembered.handle.name; node("sharedRootName").title = remembered.handle.name; }
            if (remembered && await remembered.handle.queryPermission({ mode: "readwrite" }) === "granted") {
                await WorkloadShare.connect(remembered.handle, remembered.key); node("connectShared").textContent = "共有";
            } else if (remembered) { info("前回の共有先を記憶しています。「ログ集計」で読み書きを許可してください。共有先を変更する場合は「設定」の「共有」を選択してください。"); }
            if (!window.showDirectoryPicker || !window.isSecureContext) { info("この開き方では共有フォルダを利用できません。通常のEdgeでHTMLファイルを開いてください。", true); }
        } catch (error) { info("設定を復元できません：" + (error.message || String(error)) + "　ファイルと共有先を選択してください。", true); }
    }
    return { init: init, choose: choose, chooseRoot: chooseRoot, connectShared: connectShared, retryShared: retryShared,
        selectLogs: selectLogs, runWeekday: runWeekday, sharedStatus: sharedStatus,
        selectWeekendLogs: selectWeekendLogs, runWeekend: runWeekend, isChoosing: function () { return choosing; },
        exportHistory: function () { download(WorkloadShare.historyText(), WorkloadShare.dayKey(new Date()) + "_HTML履歴.json"); },
        exportPending: function () { try { download(JSON.stringify({ schemaVersion: 1, application: "WorkloadRankingHTML", jobs: WorkloadShare.pendingExport() }, null, 2), "業務量モニター_保存待ち控え.json"); } catch (error) { fail(error); } }
    };
}());

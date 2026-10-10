/* HTML edition: calculation and validation from repository c9cb73c. */
var WorkloadImport = (function () {
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
        throw new Error("workload-config.js が見つからないか、設定の書式が不正です。HTMLと同じフォルダに配置してください。");
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
            unreadFolder: joinPath(joinPath(c.paths.unreadRoot, dateFolder(date, "yyyymm")), dateFolder(date, c.unread.folderFormat)),
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

    function openLink(url) {
        var target = validURL(url);
        if (!target) { throw new Error("リンクが未設定か、http/httpsのURLではありません。"); }
        window.open(target, "_blank", "noopener,noreferrer");
    }
    return { parseCSV: parseCSV, buildData: buildData, buildFromLogs: buildFromLogs,
        readLogRows: readLogRows, planPaths: planPaths, validateConfig: validateConfig,
        keyOf: keyOf, columnIndex: columnIndex, columnLabel: columnLabel,
        detectCharset: detectCharset, validURL: validURL, openLink: openLink };
}());

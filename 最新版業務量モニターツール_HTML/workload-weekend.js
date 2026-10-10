/*
 * 土日ログの設定。列・行が変わった場合は、下記だけを変更してください。
 * 列はExcelの列記号（A、B、C...）、行は1から数えた行番号で指定します。
 * itemsは8列です。記載した順番で右の集計項目を表示します。
 * 項目名はheaderRowの見出しから取得し、startRowから件数を集計します。
 * 変更後はUTF-8で保存し、HTMLを開き直してください。
 * 平日の設定・振分表・共有保存処理には影響しません。
 */
var WorkloadWeekendConfig = {
    headerRow: 2,                       // 項目名の見出し行
    startRow: 3,                        // データの開始行
    columns: {
        organization: "C",             // 組織（B列のエリアは照合に使いません）
        caName: "D",                   // CA名
        items: ["F", "G", "H", "I", "J", "K", "L", "M"], // 合計に含める8項目
        unsorted: "R"                  // 未仕分（合計には含めません）
    }
};

/* 土日ログの集計・対象項目。平日の計算と共有保存は既存処理を使う。 */
var WorkloadWeekend = (function () {
    var labels = defaultLabels();
    var enabled = [true, true, true, true, true, true, true, true, true], loaded = false, sourceDay = "";
    function trim(value) { return String(value == null ? "" : value).trim(); }
    function normal(value) { return trim(value).replace(/[\s\u3000]+/g, ""); }
    function defaultLabels() {
        var items = WorkloadWeekendConfig && WorkloadWeekendConfig.columns && WorkloadWeekendConfig.columns.items || [], names = [];
        for (var i = 0; i < 8; i++) { names.push(items[i] ? trim(items[i]).toUpperCase() + "列" : "項目" + (i + 1)); }
        return names.concat(["未仕分"]);
    }
    function logSettings() {
        var c = WorkloadWeekendConfig, columns = c && c.columns, positions, seen = {};
        if (!c || !Number.isInteger(c.headerRow) || c.headerRow < 1 || !Number.isInteger(c.startRow) || c.startRow <= c.headerRow) {
            throw new Error("土日ログの設定：headerRowは1以上、startRowは見出し行より後の整数で指定してください。");
        }
        if (!columns || !Array.isArray(columns.items) || columns.items.length !== 8) {
            throw new Error("土日ログの設定：columns.itemsには集計する8列を指定してください。");
        }
        positions = [columns.organization, columns.caName].concat(columns.items, [columns.unsorted]).map(function (column) { return WorkloadImport.columnIndex(column); });
        positions.forEach(function (position) {
            if (seen[position]) { throw new Error("土日ログの設定：列指定が重複しています（" + WorkloadImport.columnLabel(position) + "列）。"); }
            seen[position] = true;
        });
        return { headerRow: c.headerRow, startRow: c.startRow, organization: positions[0], caName: positions[1], values: positions.slice(2), maximum: Math.max.apply(null, positions) };
    }
    function add(a, b, context) {
        var value = a + b;
        if (!Number.isSafeInteger(value)) { throw new Error(context + "：合計件数が大きすぎます。"); }
        return value;
    }
    function log(rows, label) {
        var settings = logSettings(), header = rows[settings.headerRow - 1], lastColumn = WorkloadImport.columnLabel(settings.maximum);
        if (!header || header.length <= settings.maximum) { throw new Error(label + "：" + settings.headerRow + "行目の見出しと" + lastColumn + "列までの列を確認してください。"); }
        var names = [], result = { map: {}, order: [], rowCount: 0, duplicateCount: 0 }, counts = {}, i, j, row, key, item, context;
        for (j = 0; j < 9; j++) {
            names.push(trim(header[settings.values[j]]) || (j < 8 ? WorkloadImport.columnLabel(settings.values[j]) + "列" : "未仕分"));
        }
        if (names.some(function (name) { return /\uFFFD/.test(name); })) { throw new Error(label + "：見出しの文字化けを検出しました。"); }
        for (i = settings.startRow - 1; i < rows.length; i++) {
            row = rows[i]; if (row.every(function (value) { return !trim(value); })) { continue; }
            context = label + " " + (i + 1) + "行目";
            if (row.length <= settings.maximum) { throw new Error(context + "：" + lastColumn + "列まで必要ですが、列数が足りません。"); }
            var organization = trim(row[settings.organization]), name = trim(row[settings.caName]);
            if (!organization || !name) { throw new Error(context + "：組織またはCA名が空欄です（" + WorkloadImport.columnLabel(settings.organization) + "列・" + WorkloadImport.columnLabel(settings.caName) + "列）。"); }
            if (/\uFFFD/.test(organization + name)) { throw new Error(context + "：文字化けを検出しました。"); }
            key = WorkloadImport.keyOf(organization, name); item = result.map[key];
            if (!item) {
                item = { key: key, organization: organization, name: name, unread: 0, flag: 0, unsorted: 0 };
                result.map[key] = item; result.order.push(key); counts[key] = [0, 0, 0, 0, 0, 0, 0, 0, 0];
            } else { result.duplicateCount++; }
            for (j = 0; j < 9; j++) {
                var column = settings.values[j];
                var value = WorkloadImport.countValue(row[column], context + " " + WorkloadImport.columnLabel(column) + "列");
                counts[key][j] = add(counts[key][j], value, context);
                if (j < 8) { item.unread = add(item.unread, value, context); }
                else { item.unsorted = add(item.unsorted, value, context); }
            }
            result.rowCount++;
        }
        return { labels: names, log: result, counts: counts };
    }
    function build(logs, roster, previousStaff, config) {
        if (logs.length !== 3) { throw new Error("土日ログはCSVを3つ選択してください。"); }
        var merged = { map: {}, order: [], rowCount: 0, duplicateCount: 0 }, counts = {}, names = logs[0].labels, columns = logSettings().values, i, j, key, current, item;
        for (i = 0; i < logs.length; i++) {
            for (j = 0; j < 9; j++) {
                if (normal(names[j]) !== normal(logs[i].labels[j])) { throw new Error("土日ログ " + (i + 1) + "つ目：項目名が最初のCSVと一致しません（" + WorkloadImport.columnLabel(columns[j]) + "列）。同じ日の同じ項目のCSVを選択してください。"); }
            }
            for (var n = 0; n < logs[i].log.order.length; n++) {
                key = logs[i].log.order[n]; item = logs[i].log.map[key]; current = merged.map[key];
                if (!current) {
                    current = { key: key, organization: item.organization, name: item.name, unread: 0, flag: 0, unsorted: 0 };
                    merged.map[key] = current; merged.order.push(key); counts[key] = [0, 0, 0, 0, 0, 0, 0, 0, 0];
                } else { merged.duplicateCount++; }
                current.unread = add(current.unread, item.unread, "土日ログ"); current.unsorted = add(current.unsorted, item.unsorted, "土日ログ");
                for (j = 0; j < 9; j++) { counts[key][j] = add(counts[key][j], logs[i].counts[key][j], "土日ログ"); }
            }
            merged.rowCount += logs[i].log.rowCount; merged.duplicateCount += logs[i].log.duplicateCount;
        }
        var result = WorkloadImport.buildFromLogs(merged, { map: {}, order: [], rowCount: 0, duplicateCount: 0 }, roster.rows, previousStaff, config, roster.linksByRow);
        var overall = 0, itemTotals = [0, 0, 0, 0, 0, 0, 0, 0, 0];
        result.staffData.forEach(function (staff) {
            var staffTotal = 0;
            staff.cas.forEach(function (ca) {
                ca.weekendCounts = counts[ca.importKey] || [0, 0, 0, 0, 0, 0, 0, 0, 0];
                for (var i = 0; i < 9; i++) { itemTotals[i] = add(itemTotals[i], ca.weekendCounts[i], "全体の土日ログ"); }
                staffTotal = add(staffTotal, ca.unread, "担当者別の土日ログ");
            });
            overall = add(overall, staffTotal, "全体の土日ログ");
        });
        result.weekend = { labels: names.slice(), day: "" }; result.warnings = roster.warnings || [];
        result.hourlyRate = 10; result.mode = "weekend"; return result;
    }
    function accept(meta) {
        var same = loaded && sourceDay === meta.day && labels.every(function (name, index) { return normal(name) === normal(meta.labels[index]); });
        labels = meta.labels.slice(); sourceDay = meta.day; loaded = true;
        if (!same) { enabled = [true, true, true, true, true, true, true, true, true]; }
    }
    function applySelection(data) {
        data.forEach(function (staff) {
            staff.cas.forEach(function (ca) {
                var values = ca.weekendCounts || [0, 0, 0, 0, 0, 0, 0, 0, 0];
                ca.unread = 0; ca.flag = 0;
                for (var i = 0; i < 8; i++) { if (enabled[i]) { ca.unread += values[i]; } }
                ca.unsorted = enabled[8] ? values[8] : 0;
            });
        });
    }
    function setItem(index, checked) {
        if (!loaded || !Number.isInteger(index) || index < 0 || index > 8) { return false; }
        enabled[index] = !!checked; return true;
    }
    function render(data, excluded) {
        var totals = [0, 0, 0, 0, 0, 0, 0, 0, 0];
        data.forEach(function (staff) {
            if (excluded(staff) || staff.cas.every(function (ca) { return !!ca.committedAt; })) { return; }
            staff.cas.forEach(function (ca) {
                var values = ca.weekendCounts || [];
                for (var i = 0; i < 9; i++) { totals[i] += values[i] || 0; }
            });
        });
        for (var i = 0; i < 9; i++) {
            var check = document.getElementById("weekendCheck" + i), label = document.getElementById("weekendLabel" + i);
            if (!check || !label) { continue; }
            label.textContent = labels[i]; label.title = labels[i]; check.checked = enabled[i]; check.disabled = !loaded;
            check.setAttribute("aria-label", labels[i] + (i < 8 ? "を合計に含める" : "の未仕分表示を有効にする"));
            document.getElementById("weekendItem" + i).className = "weekend-item" + (enabled[i] ? "" : " is-excluded");
            document.getElementById("weekendValue" + i).textContent = String(totals[i]).replace(/\B(?=(\d{3})+(?!\d))/g, ",");
        }
    }
    function reset() {
        labels = defaultLabels();
        enabled = [true, true, true, true, true, true, true, true, true]; loaded = false; sourceDay = "";
    }
    return { readLog: log, build: build, accept: accept, applySelection: applySelection, setItem: setItem, render: render, reset: reset };
}());

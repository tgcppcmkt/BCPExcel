/* Browser shared history. Each operation is stored in its own immutable file. */
var WorkloadShare = (function () {
    function trim(value) { return String(value === null || typeof value === "undefined" ? "" : value).replace(/^\s+|\s+$/g, ""); }
    function normalized(value) { return trim(value).toLowerCase().replace(/[\s\u3000]+/g, ""); }
    function keyOf(organization, caName) {
        var a = trim(organization).replace(/[\s\u3000]+/g, ""), b = trim(caName).replace(/[\s\u3000]+/g, "");
        return "@" + a.length + ":" + a + "|" + b.length + ":" + b;
    }
    /* 担当者IDを別の名前空間に置き、既存の共有JSON形式・保存処理を維持する。 */
    function staffSupportKey(staff) { return keyOf("@WorkloadStaffSupport", staff.id); }
    function pad(value) { return value < 10 ? "0" + value : String(value); }
    function dayKey(date) { return String(date.getFullYear()) + pad(date.getMonth() + 1) + pad(date.getDate()); }
    function timeText(date) { return date.getFullYear() + "-" + pad(date.getMonth() + 1) + "-" + pad(date.getDate()) + " " + pad(date.getHours()) + ":" + pad(date.getMinutes()) + ":" + pad(date.getSeconds()); }
    function json(value) { return JSON.stringify(value).replace(/\u2028/g, "\\u2028").replace(/\u2029/g, "\\u2029"); }
    function own(object, name) { return Object.prototype.hasOwnProperty.call(object, name); }
    function empty(day) { return { schemaVersion: 2, date: day, supports: [], commits: [], revision: 0, ids: {}, caChecks: [], supportRequests: [] }; }
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
    function checked(state, key) {
        var i, rows = state.caChecks || [];
        for (i = 0; i < rows.length; i++) { if (keyOf(rows[i].organization, rows[i].caName) === key) { return rows[i]; } }
        return null;
    }
    function supportRequested(state, key) {
        var i, rows = state.supportRequests || [];
        for (i = 0; i < rows.length; i++) { if (keyOf(rows[i].organization, rows[i].caName) === key) { return rows[i]; } }
        return null;
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
        if (record.caChecks) {
            if (Object.prototype.toString.call(record.caChecks) !== "[object Array]") { throw new Error("共有履歴のCAチェック情報が不正です。"); }
            for (i = 0; i < record.caChecks.length; i++) {
                item = record.caChecks[i]; validateIdentity(item); validateTime(item.checkedAt);
                if (item.id !== record.id || checked(state, keyOf(item.organization, item.caName))) { throw new Error("共有履歴のCAチェックが重複しています。"); }
                state.caChecks.push({ id: item.id, organization: item.organization, caName: item.caName, checkedAt: item.checkedAt });
            }
        }
        if (record.supportRequests) {
            if (Object.prototype.toString.call(record.supportRequests) !== "[object Array]") { throw new Error("共有履歴のサポ希望情報が不正です。"); }
            for (i = 0; i < record.supportRequests.length; i++) {
                item = record.supportRequests[i]; validateIdentity(item); validateTime(item.requestedAt);
                key = keyOf(item.organization, item.caName);
                if (item.id !== record.id || supportRequested(state, key) || committed(state, key)) { throw new Error("共有履歴のサポ希望が重複しているか、コミット済みです。"); }
                state.supportRequests.push({ id: item.id, organization: item.organization, caName: item.caName, requestedAt: item.requestedAt });
            }
        }
        if (record.cancelSupportRequests) {
            if (Object.prototype.toString.call(record.cancelSupportRequests) !== "[object Array]") { throw new Error("共有履歴のサポ希望取消情報が不正です。"); }
            for (i = 0; i < record.cancelSupportRequests.length; i++) {
                item = record.cancelSupportRequests[i]; validateIdentity(item);
                found = supportRequested(state, keyOf(item.organization, item.caName));
                if (!found || found.id !== item.id) { throw new Error("取り消すサポ希望が共有履歴と一致しません。"); }
                for (j = 0; j < state.supportRequests.length; j++) { if (state.supportRequests[j] === found) { state.supportRequests.splice(j, 1); break; } }
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
        var ids = [], id, value;
        for (id in state.ids) { if (own(state.ids, id)) { ids.push(id.slice(1)); } }
        value = { schemaVersion: 2, date: state.date, supports: state.supports, commits: state.commits,
            revision: state.revision, operations: ids };
        if (state.caChecks && state.caChecks.length) { value.caChecks = state.caChecks; }
        if (state.supportRequests && state.supportRequests.length) { value.supportRequests = state.supportRequests; }
        return JSON.stringify(value, null, 2) + "\n";
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
        if (typeof value.caChecks !== "undefined") {
            if (Object.prototype.toString.call(value.caChecks) !== arrayTag) { throw new Error("共有JSONのCAチェック情報が不正です。"); }
            for (i = 0; i < value.caChecks.length; i++) {
                row = value.caChecks[i]; validateIdentity(row); validateTime(row.checkedAt);
                if (typeof row.id !== "string" || !own(ids, "@" + row.id) || checked(state, keyOf(row.organization, row.caName))) { throw new Error("共有JSONのCAチェックが不正です。"); }
                state.caChecks.push({ id: row.id, organization: row.organization, caName: row.caName, checkedAt: row.checkedAt });
            }
        }
        if (typeof value.supportRequests !== "undefined") {
            if (Object.prototype.toString.call(value.supportRequests) !== arrayTag) { throw new Error("共有JSONのサポ希望情報が不正です。"); }
            for (i = 0; i < value.supportRequests.length; i++) {
                row = value.supportRequests[i]; validateIdentity(row); validateTime(row.requestedAt);
                if (typeof row.id !== "string" || !own(ids, "@" + row.id) || supportRequested(state, keyOf(row.organization, row.caName))) { throw new Error("共有JSONのサポ希望が不正です。"); }
                state.supportRequests.push({ id: row.id, organization: row.organization, caName: row.caName, requestedAt: row.requestedAt });
            }
        }
        state.ids = ids; state.revision = value.revision;
        return state;
    }
    function makeRecord(state, action, id, now) {
        var record = { id: id, supports: [], commits: [] }, key = keyOf(action.organization, action.caName), list = active(state, key), i, row, name = trim(action.supportName), stamp = timeText(now);
        validateIdentity(action);
        if (own(state.ids, "@" + id)) { return record; }
        if (action.kind === "staffSupportRequest" && action.requested === false) {
            if (Object.prototype.toString.call(action.legacyRequests) !== "[object Array]") { throw new Error("担当者のサポ希望情報が不正です。"); }
            record.cancelSupportRequests = [];
            for (i = 0; i < action.legacyRequests.length; i++) {
                row = action.legacyRequests[i]; validateIdentity(row);
                list = supportRequested(state, keyOf(row.organization, row.caName));
                if (!list) { continue; }
                if (list.id !== row.requestId) { throw new Error("サポ希望が更新されています。最新の状態を確認してから、もう一度取り消してください。"); }
                record.cancelSupportRequests.push({ id: list.id, organization: list.organization, caName: list.caName });
            }
        }
        if (action.kind === "supportRequest" || action.kind === "staffSupportRequest") {
            if (typeof action.requested !== "boolean") { throw new Error("サポ希望の指定が不正です。"); }
            row = supportRequested(state, key);
            if (action.requested) {
                if (committed(state, key)) { throw new Error("このCAはコミット済みです。先にコミットを取り消してください。"); }
                if (row) { return record; }
                record.supportRequests = [{ id: id, organization: trim(action.organization), caName: trim(action.caName), requestedAt: stamp }];
            } else {
                if (!row) { return record; }
                /* 古い画面から、他のフォームが新しく設定した希望を取り消さない。 */
                if (row.id !== action.requestId) { throw new Error("サポ希望が更新されています。最新の状態を確認してから、もう一度取り消してください。"); }
                record.cancelSupportRequests = record.cancelSupportRequests || [];
                record.cancelSupportRequests.push({ id: row.id, organization: row.organization, caName: row.caName });
            }
        } else if (action.kind === "check") {
            if (checked(state, key)) { return record; }
            if (committed(state, key)) { throw new Error("このCAはコミット済みです。先にコミットを取り消してください。"); }
            record.caChecks = [{ id: id, organization: trim(action.organization), caName: trim(action.caName), checkedAt: stamp }];
        } else if (action.kind === "commit") {
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
    var root = null, listener = null, running = false, timer = null, busy = false, lastState = null, lastSignature = "", lastError = "";
    var contexts = new Map(), callbacks = new Map(), rootHandles = new Map();
    var pendingPrefix = "workload-html.pending.v1.", dataDirectory = "WORKLOAD_HTML_DATA";
    var kinds = ["check", "supportRequest", "staffSupportRequest", "commit", "cancelCommit", "add", "finish"];
    function clone(value) { return JSON.parse(JSON.stringify(value)); }
    function failure(message, terminal) { var error = new Error(message); error.terminal = !!terminal; return error; }
    function uid() {
        var bytes = new Uint8Array(16); window.crypto.getRandomValues(bytes);
        return "h-" + Array.from(bytes, function (b) { return ("0" + b.toString(16)).slice(-2); }).join("");
    }
    function pendingJobs() {
        var result = [], storage = window.localStorage, i, key, job;
        for (i = 0; i < storage.length; i++) {
            key = storage.key(i);
            if (key && key.indexOf(pendingPrefix) === 0) {
                try { job = JSON.parse(storage.getItem(key)); }
                catch (error) { throw failure("端末内の保存待ちデータが壊れています。控えを保全してから管理担当へ確認してください。"); }
                if (!job || key !== pendingPrefix + job.id || !/^h-[a-f0-9]{32}$/.test(job.id) ||
                    !/^\d{8}$/.test(job.date) || typeof job.rootKey !== "string" || !job.action || !job.stamp) {
                    throw failure("端末内の保存待ちデータの形式が不正です。自動的には削除しません。");
                }
                result.push(job);
            }
        }
        return result.sort(function (a, b) { return a.stamp.localeCompare(b.stamp) || a.id.localeCompare(b.id); });
    }
    function keepJob(job) {
        try { window.localStorage.setItem(pendingPrefix + job.id, JSON.stringify(job)); }
        catch (error) { throw failure("保存待ちをこの端末に退避できません。ブラウザの保存設定・空き容量を確認してください。操作は確定していません。"); }
    }
    function removeJob(job) { window.localStorage.removeItem(pendingPrefix + job.id); }
    async function optionalFile(folder, name) {
        try { return await folder.getFileHandle(name); }
        catch (error) { if (error.name === "NotFoundError") { return null; } throw error; }
    }
    async function optionalDirectory(folder, name) {
        try { return await folder.getDirectoryHandle(name); }
        catch (error) { if (error.name === "NotFoundError") { return null; } throw error; }
    }
    async function belongs(job) {
        if (!root) { return false; }
        if (job.rootKey === root.key) { return true; }
        if (!rootHandles.has(job.rootKey)) { rootHandles.set(job.rootKey, await WorkloadBrowserStore.get(job.rootKey)); }
        var handle = rootHandles.get(job.rootKey);
        return !!handle && await root.handle.isSameEntry(handle);
    }
    async function permission() {
        if (!root || await root.handle.queryPermission({ mode: "readwrite" }) !== "granted") {
            throw failure("共有フォルダへの読み書きが許可されていません。「共有フォルダを接続」または「保存を再試行」で許可してください。");
        }
    }
    async function dayContext(day) {
        if (contexts.has(day)) { return contexts.get(day); }
        var base = empty(day), legacyName = day + ".json", fileHandle = await optionalFile(root.handle, legacyName), file;
        if (fileHandle) {
            file = await fileHandle.getFile();
            if (file.size > 16 * 1024 * 1024) { throw failure("移行元の日付JSONが大きすぎます。"); }
            base = parse(await file.text(), day);
        } else {
            fileHandle = await optionalFile(root.handle, day + ".js");
            if (fileHandle) {
                file = await fileHandle.getFile();
                if (file.size > 16 * 1024 * 1024) { throw failure("移行元の日付履歴が大きすぎます。"); }
                base = parseLegacy(await file.text(), day);
            }
        }
        var ctx = { day: day, base: base, baseHandle: fileHandle,
            baseMark: file ? file.size + ":" + file.lastModified : "", events: new Map(),
            order: [], state: clone(base), rejected: [], maximum: 0, snapshot: clone(base), signature: "", integrityCursor: 0, invalidError: null };
        contexts.set(day, ctx); return ctx;
    }
    function validateEvent(event, name, day) {
        if (!event || event.schemaVersion !== 1 || event.application !== "WorkloadRankingHTML" ||
            !/^h-[a-f0-9]{32}$/.test(event.id) || name !== "op-" + event.id + ".json" || event.date !== day ||
            !Number.isSafeInteger(event.clock) || event.clock < 1 || event.clock > 1000000000000 ||
            typeof event.stamp !== "string" || !/^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d\.\d{3}Z$/.test(event.stamp) ||
            !Number.isFinite(Date.parse(event.stamp)) || dayKey(new Date(event.stamp)) !== day ||
            !event.action || kinds.indexOf(event.action.kind) < 0) {
            throw failure("共有履歴の形式・日付・操作番号が不正です：" + name);
        }
        validateIdentity(event.action);
    }
    async function scanDay(day) {
        var ctx = await dayContext(day), storage = await optionalDirectory(root.handle, dataDirectory);
        if (ctx.invalidError) { throw ctx.invalidError; }
        var folder = storage ? await optionalDirectory(storage, day) : null, seen = new Set(), count = 0;
        // Immutable operation bodies stay cached. Check existing metadata in small,
        // rotating batches so a busy network share is not opened thousands of times per poll.
        var cachedNames = Array.from(ctx.events.keys()), checks = new Set();
        for (var cursor = 0; cursor < Math.min(16, cachedNames.length); cursor++) { checks.add(cachedNames[(ctx.integrityCursor + cursor) % cachedNames.length]); }
        ctx.integrityCursor = cachedNames.length ? (ctx.integrityCursor + checks.size) % cachedNames.length : 0;
        var currentBase = await optionalFile(root.handle, day + ".json") || await optionalFile(root.handle, day + ".js");
        if (!!currentBase !== !!ctx.baseHandle || currentBase && !await currentBase.isSameEntry(ctx.baseHandle)) {
            throw failure("移行元の日付履歴が新規作成・置換されています。旧版を停止し、全員がHTML版を開き直してください。");
        }
        if (ctx.baseHandle) {
            var baseFile = await ctx.baseHandle.getFile();
            if (ctx.baseMark !== baseFile.size + ":" + baseFile.lastModified) {
                throw failure("移行元の日付履歴が変更されています。HTML版の使用中は旧版から更新しないでください。");
            }
        }
        if (folder) {
            for await (var pair of folder.entries()) {
                var name = pair[0], handle = pair[1];
                if (handle.kind !== "file" || !/^op-h-[a-f0-9]{32}\.json$/.test(name)) { continue; }
                if (++count > 50000) { throw failure("当日の共有操作が50,000件を超えています。履歴を保全して管理担当へ確認してください。"); }
                seen.add(name);
                var cached = ctx.events.get(name);
                if (cached && !checks.has(name)) { continue; }
                var file = await handle.getFile();
                // A newly created event is empty until its writable stream closes.
                if (!file.size) { if (cached) { ctx.invalidError = failure("保存済みの共有履歴が空になっています：" + name); throw ctx.invalidError; } continue; }
                var mark = file.size + ":" + file.lastModified;
                if (cached) {
                    if (cached.mark !== mark) { ctx.invalidError = failure("保存済みの共有履歴が書き換えられています：" + name); throw ctx.invalidError; }
                    continue;
                }
                if (file.size > 2 * 1024 * 1024) { throw failure("共有操作のファイルが大きすぎます：" + name); }
                var event;
                try { event = JSON.parse(await file.text()); }
                catch (error) { throw failure("共有履歴を読み込めません。元ファイルは変更しません：" + name); }
                validateEvent(event, name, day); ctx.events.set(name, { event: event, mark: mark });
            }
        }
        for (var cachedName of ctx.events.keys()) {
            if (!seen.has(cachedName)) { throw failure("保存済みの共有履歴が削除されています：" + cachedName); }
        }
        var ordered = Array.from(ctx.events.values(), function (x) { return x.event; });
        ordered.sort(function (a, b) { return a.clock - b.clock || a.id.localeCompare(b.id); });
        var order = ordered.map(function (x) { return x.id; }), prefix = ctx.order.length <= order.length;
        for (var i = 0; prefix && i < ctx.order.length; i++) { if (ctx.order[i] !== order[i]) { prefix = false; } }
        if (!prefix) { ctx.state = clone(ctx.base); ctx.rejected = []; ctx.order = []; }
        for (i = ctx.order.length; i < ordered.length; i++) {
            var row = ordered[i];
            try { apply(ctx.state, makeRecord(ctx.state, row.action, row.id, new Date(row.stamp))); }
            catch (error) {
                ctx.rejected.push({ id: row.id, caName: row.action.caName, organization: row.action.organization,
                    kind: row.action.kind, stamp: timeText(new Date(row.stamp)), message: error.message || String(error) });
            }
        }
        ctx.order = order; ctx.maximum = ordered.length ? ordered[ordered.length - 1].clock : 0;
        var signature = day + ":" + order.join(",");
        if (ctx.signature !== signature) { ctx.snapshot = clone(ctx.state); ctx.signature = signature; }
        return ctx;
    }
    function publish(ctx, error) {
        lastState = ctx.snapshot;
        var message = error ? error.message || String(error) : "";
        if (listener && (lastError !== message || error || lastSignature !== ctx.signature)) {
            lastSignature = ctx.signature; listener(error || null, lastState);
        }
        lastError = message; status(ctx, error);
    }
    function status(ctx, error) {
        if (typeof WorkloadBrowser !== "undefined") {
            var count = 0;
            try { count = pendingJobs().length; } catch (badPending) { error = badPending; }
            WorkloadBrowser.sharedStatus({ connected: !!root, folder: root ? root.handle.name : "",
                pending: count, rejected: ctx ? ctx.rejected : [], error: error ? error.message : "" });
        }
    }
    function report(error) { lastError = error.message || String(error); if (listener) { listener(error, lastState); } status(null, error); }
    async function writeEvent(job, event) {
        var storage = await root.handle.getDirectoryHandle(dataDirectory, { create: true });
        var folder = await storage.getDirectoryHandle(job.date, { create: true });
        var name = "op-" + job.id + ".json", text = JSON.stringify(event, null, 2) + "\n";
        var handle = await optionalFile(folder, name);
        if (handle) {
            var existing = await handle.getFile();
            if (existing.size) {
                if (await existing.text() !== text) { throw failure("同じ操作番号の履歴が別の内容で保存されています。上書きしません。", true); }
                return;
            }
        } else { handle = await folder.getFileHandle(name, { create: true }); }
        var stream = await handle.createWritable();
        try { await stream.write(text); await stream.close(); }
        catch (error) { try { await stream.abort(); } catch (ignored) {} throw error; }
        var written = await handle.getFile();
        if (await written.text() !== text) { throw failure("共有履歴の保存結果を確認できません。端末内の控えから再試行します。"); }
    }
    async function drain() {
        if (!running || !root || busy) { return; }
        busy = true;
        try {
            await permission();
            var jobs = pendingJobs(), failed = null;
            for (var i = 0; i < jobs.length; i++) {
                var job = jobs[i]; if (!await belongs(job)) { continue; }
                var callback = callbacks.get(job.id), error = null;
                try {
                    var ctx = await scanDay(job.date);
                    if (!own(ctx.state.ids, "@" + job.id)) {
                        if (!job.event) {
                            try { makeRecord(ctx.state, job.action, job.id, new Date(job.stamp)); }
                            catch (badAction) { badAction.terminal = true; throw badAction; }
                            job.event = { schemaVersion: 1, application: "WorkloadRankingHTML", id: job.id, date: job.date,
                                clock: ctx.maximum + 1, stamp: job.stamp, action: job.action };
                            keepJob(job);
                        }
                        validateEvent(job.event, "op-" + job.id + ".json", job.date);
                        await writeEvent(job, job.event);
                        ctx = await scanDay(job.date);
                        var rejection = ctx.rejected.find(function (x) { return x.id === job.id; });
                        if (rejection) { throw failure("同時操作により反映できませんでした。最新の状態を確認してください。\n" + rejection.message, true); }
                        if (!own(ctx.state.ids, "@" + job.id)) { throw failure("共有側で操作の反映を確認できません。控えを残して再試行します。"); }
                    }
                    removeJob(job);
                } catch (saveError) {
                    error = saveError;
                    if (error.terminal) { removeJob(job); }
                }
                var today;
                try { today = await scanDay(dayKey(new Date())); }
                catch (readError) { if (!error) { error = readError; } today = { snapshot: lastState, signature: lastSignature, rejected: [] }; }
                callbacks.delete(job.id);
                publish(today, error);
                if (callback) { callback(error, lastState); }
                else if (!error && typeof WorkloadUI !== "undefined") { WorkloadUI.browserWriteRecovered(job.action, lastState); }
                if (error && !error.terminal) { failed = error; break; }
            }
            publish(await scanDay(dayKey(new Date())), failed);
        } catch (error) {
            for (var entry of callbacks) { entry[1](error, lastState); }
            callbacks.clear(); report(error);
        } finally { busy = false; }
    }
    async function refresh() {
        if (!running || !root || busy) { return; }
        await drain();
    }
    function request(action, callback) {
        if (!root || !running) { callback(failure("先に全員が使う「共有フォルダを接続」してください。"), lastState); return; }
        // Retrying a confirmed operation reuses its durable identity.
        var pending;
        try { pending = pendingJobs().find(function (x) { return x.rootKey === root.key && x.date === dayKey(new Date()) && JSON.stringify(x.action) === JSON.stringify(action); }); }
        catch (error) { callback(error, lastState); return; }
        if (pending) { callbacks.set(pending.id, callback); drain(); return; }
        var job = { id: uid(), date: dayKey(new Date()), stamp: new Date().toISOString(), action: clone(action),
            rootKey: root.key, rootName: root.handle.name };
        try { keepJob(job); } catch (error) { callback(error, lastState); return; }
        callbacks.set(job.id, callback); status(null); drain();
    }
    function start(callback) {
        stop(); listener = callback; running = true; lastSignature = "";
        if (root) { refresh(); }
        else { listener(failure("全員が使う共有フォルダを接続してください。"), empty(dayKey(new Date()))); }
        timer = window.setInterval(function () { refresh().catch(report); }, 3000);
    }
    function stop() { running = false; if (timer) { window.clearInterval(timer); timer = null; } }
    async function connect(handle, key) {
        if (busy) { throw failure("保存・読込が完了してから接続してください。"); }
        var previous = root;
        if (previous && !await previous.handle.isSameEntry(handle)) {
            for (var job of pendingJobs()) { if (await belongs(job)) { throw failure("現在の共有先への保存待ちがあります。保存完了後に共有先を変更してください。"); } }
        }
        if (await handle.queryPermission({ mode: "readwrite" }) !== "granted") { throw failure("共有フォルダへの読み書きの許可が必要です。"); }
        // Store the actual handle before accepting writes; a folder name alone is not an identity.
        await WorkloadBrowserStore.set(key, handle);
        root = { handle: handle, key: key }; rootHandles.set(key, handle);
        contexts = new Map(); lastSignature = "";
        try {
            await handle.getDirectoryHandle(dataDirectory, { create: true });
            var ctx = await scanDay(dayKey(new Date())); publish(ctx, null);
        } catch (error) { root = previous; contexts = new Map(); lastSignature = ""; report(error); throw error; }
        if (typeof WorkloadUI !== "undefined") { WorkloadUI.refreshBrowserShare(); }
        await drain();
    }
    return {
        start: start, refresh: refresh, stop: stop, connect: connect,
        isConnected: function () { return !!root; }, isBusy: function () { return busy; },
        hasPendingWrites: function () { try { return callbacks.size > 0 || pendingJobs().length > 0; } catch (error) { return true; } },
        pendingExport: function () { return pendingJobs(); },
        snapshot: function () { return lastState ? clone(lastState) : empty(dayKey(new Date())); },
        historyText: function () { return encode(lastState || empty(dayKey(new Date()))); },
        checkCA: function (ca, callback) { request({ kind: "check", organization: ca.organization, caName: ca.name }, callback); },
        setSupportRequest: function (ca, requested, callback) { request({ kind: "supportRequest", organization: ca.organization, caName: ca.name, requested: requested, requestId: ca.supportRequestId || "" }, callback); },
        staffSupportKey: staffSupportKey,
        setStaffSupportRequest: function (staff, requested, callback) {
            var rows = [], i, ca;
            for (i = 0; i < staff.cas.length; i++) { ca = staff.cas[i]; rows.push({ organization: ca.organization, caName: ca.name, requestId: ca.supportRequestId || "" }); }
            request({ kind: "staffSupportRequest", organization: "@WorkloadStaffSupport", caName: staff.id,
                requested: requested, requestId: staff.supportRequestId || "", legacyRequests: rows }, callback);
        },
        addSupport: function (ca, name, callback) { request({ kind: "add", organization: ca.organization, caName: ca.name, supportName: name }, callback); },
        finishSupport: function (ca, entryId, callback) { request({ kind: "finish", organization: ca.organization, caName: ca.name, entryId: entryId }, callback); },
        commitCA: function (ca, callback) { request({ kind: "commit", organization: ca.organization, caName: ca.name }, callback); },
        cancelCommitCA: function (ca, revision, callback) { request({ kind: "cancelCommit", organization: ca.organization, caName: ca.name, committedAt: ca.committedAt, revision: revision }, callback); },
        keyOf: keyOf, dayKey: dayKey, activeForCA: active, committedForCA: committed, checkedForCA: checked,
        parseDay: parse, headerForDay: function (day) { return encode(empty(day)); }, parseLegacyDay: parseLegacy
    };

}());

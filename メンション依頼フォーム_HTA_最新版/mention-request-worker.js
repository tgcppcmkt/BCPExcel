(function () {
    var FOR_READING = 1;
    var FOR_APPENDING = 8;
    var TRISTATE_USE_DEFAULT = -2;

    var RETRY_MIN_MS = 450;
    var RETRY_MAX_MS = 1250;
    var MAX_RETRY_MS = 300000;
    var STALE_LOCK_MS = 600000;
    var PENDING_FOLDER_NAME = "MentionRequest_PerSend_Pending";
    var OUTPUT_RECEIPT_SUFFIX = ".publish";
    var MAX_OUTPUT_SEQUENCE = 9999;
    var COMPLETION_READY_SUFFIX = ".notify-ready";
    var COMPLETION_READY_MAX_MS = 30000;
    var COMPLETION_READY_POLL_MS = 200;
    var FAILURE_READY_RECENT_MS = 60000;

    // Current schema: proxy CA supports up to six names; RequestID remains last.
    var HEADER_LATEST = [
        "\u9001\u4fe1\u65e5\u6642",
        "\u62e0\u70b9",
        "\u4f9d\u983c\u8005",
        "\u4f9d\u983cNo.",
        "\u7d44\u7e54",
        "CA\u540d",
        "\u51fa\u793e\u72b6\u6cc1",
        "\u4ee3\u7406CA1",
        "\u4ee3\u7406CA2",
        "\u4ee3\u7406CA3",
        "\u4ee3\u7406CA4",
        "\u4ee3\u7406CA5",
        "\u4ee3\u7406CA6",
        "\u30aa\u30d7\u30b7\u30e7\u30f3",
        "\u30e1\u30fc\u30eb\u30e1\u30e2",
        "\u51e6\u7406\u65e5\u6642",
        "\u671f\u65e5\u307e\u305f\u306f\u9762\u63a5\u65e5\u7a0b",
        "\u30bf\u30a4\u30d7",
        "RequestID"
    ];

    // Eight-name schema; keep the six-name schema unchanged for existing Pending files.
    var HEADER_PROXY8 = HEADER_LATEST.slice(0, 13).concat(
        ["\u4ee3\u7406CA7", "\u4ee3\u7406CA8"], HEADER_LATEST.slice(13)
    );

    // Previous v28.23 schema retained so already-created 16-column Pending files can replay.
    var HEADER_V2823 = [
        "\u9001\u4fe1\u65e5\u6642",
        "\u62e0\u70b9",
        "\u4f9d\u983c\u8005",
        "\u4f9d\u983cNo.",
        "\u7d44\u7e54",
        "CA\u540d",
        "\u51fa\u793e\u72b6\u6cc1",
        "\u4ee3\u7406CA1",
        "\u4ee3\u7406CA2",
        "\u4ee3\u7406CA3",
        "\u30aa\u30d7\u30b7\u30e7\u30f3",
        "\u30e1\u30fc\u30eb\u30e1\u30e2",
        "\u51e6\u7406\u65e5\u6642",
        "\u671f\u65e5\u307e\u305f\u306f\u9762\u63a5\u65e5\u7a0b",
        "\u30bf\u30a4\u30d7",
        "RequestID"
    ];

    // v28.22 schema kept for replaying already-created Pending files.
    var HEADER_V2822 = [
        "\u9001\u4fe1\u65e5\u6642",
        "RequestID",
        "\u62e0\u70b9",
        "\u4f9d\u983c\u8005",
        "\u4f9d\u983c\u756a\u53f7",
        "\u7d44\u7e54",
        "CA\u540d",
        "\u51fa\u793e\u72b6\u6cc1",
        "\u4ee3\u7406CA1",
        "\u4ee3\u7406CA2",
        "\u4ee3\u7406CA3",
        "\u30aa\u30d7\u30b7\u30e7\u30f3",
        "\u30e1\u30fc\u30eb\u30e1\u30e2",
        "\u51e6\u7406\u65e5\u6642",
        "\u671f\u65e5\u307e\u305f\u306f\u9762\u63a5\u65e5\u7a0b",
        "\u30bf\u30a4\u30d7"
    ];

    var HEADER_LEGACY = [
        "\u9001\u4fe1\u65e5\u6642",
        "RequestID",
        "\u62e0\u70b9",
        "\u4f9d\u983c\u8005",
        "\u4f9d\u983c\u756a\u53f7",
        "\u7d44\u7e54",
        "CA\u540d",
        "\u4ee3\u7406CA\u7d44\u7e54",
        "\u4ee3\u7406CA1",
        "\u4ee3\u7406CA2",
        "\u4ee3\u7406CA3",
        "\u30e1\u30fc\u30eb\u30e1\u30e2",
        "\u51e6\u7406\u65e5\u6642",
        "\u671f\u65e5",
        "\u6642\u77ed",
        "\u81f3\u6025",
        "\u30bf\u30a4\u30d7"
    ];

    var args = WScript.Arguments;
    var fso = new ActiveXObject("Scripting.FileSystemObject");
    var shell = new ActiveXObject("WScript.Shell");

    if (args.length < 2) {
        WScript.Quit(2);
    }

    var pendingPath = String(args.Item(0) || "");
    var csvFolder = String(args.Item(1) || "");
    var notBeforeMs = 0;

    if (args.length >= 3) {
        notBeforeMs = parseInt(String(args.Item(2) || "0"), 10);
        if (isNaN(notBeforeMs) || notBeforeMs < 0) {
            notBeforeMs = 0;
        }
    }

    if (!pendingPath) {
        WScript.Quit(0);
    }

    if (!csvFolder || !fso.FolderExists(csvFolder)) {
        showFailurePopupWhenReady(pendingPath, "folder", null);
        WScript.Quit(3);
    }

    // Only retry workers for the same Pending share a lock. Different
    // submissions do not lock or append to any shared CSV.
    var lockPath = pendingPath + ".jlock";
    var receiptPath = pendingPath + OUTPUT_RECEIPT_SUFFIX;
    var startedAt = new Date().getTime();
    var pending = null;
    var outputStem = "";
    var failureStep = "lock";
    var failureError = null;

    while ((new Date().getTime() - startedAt) < MAX_RETRY_MS) {
        if (!fso.FileExists(pendingPath) && !fso.FileExists(recoveryPathFromPending(pendingPath))) {
            WScript.Quit(0);
        }

        if (!tryAcquireLock(lockPath)) {
            sleepRandom();
            continue;
        }

        if (!fso.FileExists(pendingPath) && !restorePrimaryFromRecovery(pendingPath)) {
            releaseLock(lockPath);
            WScript.Quit(0);
        }

        if (!pending) {
            try {
                failureStep = "pending";
                pending = readPendingCsv(pendingPath);
                failureStep = "name";
                outputStem = buildOutputStem(pending.requester, pending.sentAt, pending.requestId);
            } catch (ePending) {
                releaseLock(lockPath);
                showFailurePopupWhenReady(pendingPath, failureStep, ePending);
                WScript.Quit(4);
            }
        }

        try {
            publishSubmission(csvFolder, outputStem, receiptPath, pendingPath, pending.text);
        } catch (ePublish) {
            failureError = ePublish;
            releaseLock(lockPath);
            if ((new Date().getTime() - startedAt) >= MAX_RETRY_MS) {
                break;
            }
            sleepRandom();
            continue;
        }

        // Keep the receipt if Pending deletion fails, so retries can recognize
        // delivery even if the final CSV has already been collected.
        if (deletePendingPackageWithRetry(pendingPath)) {
            deleteFileWithRetry(receiptPath);
            releaseLock(lockPath);
            showSuccessPopupWhenReady(pendingPath);
            WScript.Quit(0);
        }

        releaseLock(lockPath);
        showFailurePopupWhenReady(pendingPath, "cleanup", null);
        WScript.Quit(5);
    }

    showFailurePopupWhenReady(pendingPath, failureStep, failureError);
    WScript.Quit(6);


    function readPendingCsv(path) {
        var primaryError = null;

        try {
            return readPendingCsvCore(path);
        } catch (ePrimary) {
            primaryError = ePrimary;
        }

        /*
          v28.31:
          If the Primary Pending is structurally incomplete/corrupt, validate the
          recovery sidecar first. Only a valid recovery copy may overwrite Primary.
        */
        var recoveryPath = recoveryPathFromPending(path);

        if (!fso.FileExists(recoveryPath)) {
            throw primaryError;
        }

        try {
            readPendingCsvCore(recoveryPath);
        } catch (eRecovery) {
            throw new Error("PENDING_AND_RECOVERY_INVALID");
        }

        if (!copyRecoveryToPrimary(recoveryPath, path)) {
            throw new Error("PENDING_SELF_REPAIR_FAILED");
        }

        return readPendingCsvCore(path);
    }


    function readPendingCsvCore(path) {
        var text = readText(path);
        var rows = parseCsv(text);

        if (rows.length < 2) {
            throw new Error("PENDING_EMPTY");
        }

        var schema = detectHeader(rows[0]);
        var expectedColumns = expectedColumnCount(schema);
        var expectedRequestCount = expectedRequestCountFromPendingPath(path);

        var records = [];
        var baseName = "";
        var sentAt = "";
        var requester = "";
        var requestNos = {};
        var requestIds = {};
        var i, row, no, rowRequestId, rowBaseName, rowSentAt, rowRequester;

        for (i = 1; i < rows.length; i++) {
            row = rows[i];

            if (isEmptyRecord(row)) {
                continue;
            }

            if (row.length !== expectedColumns) {
                throw new Error("PENDING_BAD_COLUMNS");
            }

            rowSentAt = sentAtFromRow(row, schema);
            rowBaseName = baseNameFromRow(row, schema);
            rowRequester = requesterFromRow(row, schema);
            rowRequestId = requestIdFromRow(row, schema);
            no = requestNoFromRow(row, schema);

            if (!rowSentAt || !rowBaseName || !rowRequestId || !no) {
                throw new Error("PENDING_KEY_MISSING");
            }

            if (!sentAt) {
                sentAt = rowSentAt;
                baseName = rowBaseName;
                requester = rowRequester;
            } else if (rowSentAt !== sentAt || rowBaseName !== baseName || rowRequester !== requester) {
                throw new Error("PENDING_MIXED_PACKAGE");
            }

            if (requestNos[no]) {
                throw new Error("PENDING_DUPLICATE_REQUEST_NO");
            }
            if (requestIds[rowRequestId]) {
                throw new Error("PENDING_DUPLICATE_REQUEST_ID");
            }

            requestNos[no] = true;
            requestIds[rowRequestId] = true;
            records.push(row);
        }

        if (records.length < 1) {
            throw new Error("PENDING_NO_DATA");
        }

        // v28.30: New Pending files carry the original request count in the
        // filename (__N1 .. __N3). Refuse to write if rows were partially lost.
        // Old Pending files have no marker and keep the legacy replay behavior.
        if (expectedRequestCount > 0) {
            if (records.length !== expectedRequestCount) {
                throw new Error("PENDING_REQUEST_COUNT_MISMATCH");
            }

            for (i = 1; i <= expectedRequestCount; i++) {
                if (!requestNos[String(i)]) {
                    throw new Error("PENDING_REQUEST_SEQUENCE_MISMATCH");
                }
            }
        }

        return {
            sentAt: sentAt,
            requester: requester,
            requestId: requestIdFromRow(records[0], schema),
            text: text
        };
    }

    function recoveryPathFromPending(path) {
        return String(path || "") + ".recovery";
    }


    function restorePrimaryFromRecovery(path) {
        var recoveryPath = recoveryPathFromPending(path);

        if (!fso.FileExists(recoveryPath)) {
            return false;
        }

        try {
            // Validate recovery before recreating Primary.
            readPendingCsvCore(recoveryPath);
        } catch (e) {
            return false;
        }

        return copyRecoveryToPrimary(recoveryPath, path);
    }


    function copyRecoveryToPrimary(recoveryPath, primaryPath) {
        var i;

        for (i = 0; i < 5; i++) {
            try {
                if (fso.FileExists(primaryPath)) {
                    fso.DeleteFile(primaryPath, true);
                }

                fso.CopyFile(recoveryPath, primaryPath, false);

                if (fso.FileExists(primaryPath)) {
                    return true;
                }
            } catch (e) {
            }

            WScript.Sleep(150 + Math.floor(Math.random() * 350));
        }

        return fso.FileExists(primaryPath);
    }


    function expectedRequestCountFromPendingPath(path) {
        var name = String(fso.GetFileName(path) || "");
        var match = /__N([1-3])(?:_|\.|$)/i.exec(name);

        if (!match) {
            return 0;
        }

        var count = parseInt(match[1], 10);
        if (isNaN(count) || count < 1 || count > 3) {
            return 0;
        }

        return count;
    }

    function publishSubmission(folder, stem, receiptPath, sourcePath, text) {
        var sequence = 1;
        var targetPath;

        // For a retry, verify the previous destination before picking another
        // name. Normal sends only check whether candidate names exist.
        failureStep = "receipt";
        // A failed receipt creation can leave a zero-byte file. No output name
        // was recorded in it, so retry the normal unused-name selection.
        if (fso.FileExists(receiptPath) && fso.GetFile(receiptPath).Size > 0) {
            var receipt = readText(receiptPath);
            var attempt = /^([1-9][0-9]{0,3})(?:\r?\n|$)/.exec(receipt);
            if (!attempt) {
                throw new Error("PUBLISH_RECEIPT_INVALID");
            }
            if (/(?:^|\r?\n)sent\r?\n/.test(receipt)) {
                return;
            }

            sequence = parseInt(attempt[1], 10);
            targetPath = submissionPath(folder, stem, sequence);
            if (fso.FileExists(targetPath) && readText(targetPath) === text) {
                markSubmissionPublished(receiptPath);
                return;
            }
        }

        // Stage in the destination folder, then publish by a same-folder
        // rename. No *.csv is visible before the stream closes and is verified.
        failureStep = "stage";
        var tempPath = fso.BuildPath(folder, "~mr-" + fso.GetBaseName(fso.GetTempName()) + ".tmp");
        var ownsTemp = false;

        try {
            // Claim our own .tmp without replacing another worker's file.
            // Copy the original bytes into this slot with no text re-encoding.
            var reservation = fso.CreateTextFile(tempPath, false, false);
            ownsTemp = true;
            reservation.Close();
            failureStep = "copy";
            fso.CopyFile(sourcePath, tempPath, true);

            failureStep = "verify";
            if (readText(tempPath) !== text) {
                throw new Error("STAGED_CSV_VERIFY_FAILED");
            }

            for (; sequence <= MAX_OUTPUT_SEQUENCE; sequence++) {
                failureStep = "name";
                targetPath = submissionPath(folder, stem, sequence);
                if (fso.FileExists(targetPath) || fso.FolderExists(targetPath)) {
                    continue;
                }

                // Record the attempted name before publication. If the move
                // succeeds but cleanup is interrupted, retry uses this name.
                failureStep = "receipt";
                var receiptStream = fso.CreateTextFile(receiptPath, true, false);
                try {
                    receiptStream.WriteLine(String(sequence));
                } finally {
                    receiptStream.Close();
                }

                failureStep = "rename";
                try {
                    // MoveFile refuses to replace an existing destination.
                    fso.MoveFile(tempPath, targetPath);
                } catch (eMove) {
                    if (fso.FileExists(tempPath) &&
                        ((eMove.number & 65535) === 58 ||
                        fso.FileExists(targetPath) || fso.FolderExists(targetPath))) {
                        continue;
                    }
                    throw eMove;
                }

                ownsTemp = false;
                markSubmissionPublished(receiptPath);
                return;
            }

            failureStep = "name";
            throw new Error("OUTPUT_NAME_LIMIT");
        } finally {
            if (ownsTemp) {
                try {
                    if (fso.FileExists(tempPath)) {
                        fso.DeleteFile(tempPath, false);
                    }
                } catch (eTemp) {
                }
            }
        }
    }


    function markSubmissionPublished(path) {
        try {
            var stream = fso.OpenTextFile(path, FOR_APPENDING, false, TRISTATE_USE_DEFAULT);
            try {
                stream.Write("\r\nsent\r\n");
            } finally {
                stream.Close();
            }
        } catch (eReceipt) {
            // A receipt write failure does not undo a verified publication.
            // A retry can still verify the recorded destination.
        }
    }


    function submissionPath(folder, stem, sequence) {
        var suffix = sequence === 1 ? "" : "_" + (sequence < 10 ? "0" : "") + sequence;
        return fso.BuildPath(folder, stem + suffix + ".csv");
    }


    function tryAcquireLock(path) {
        clearStaleLock(path);

        try {
            fso.CreateFolder(path);
            return true;
        } catch (e) {
            failureStep = "lock";
            failureError = e;
            return false;
        }
    }


    function releaseLock(path) {
        try {
            if (fso.FolderExists(path)) {
                fso.DeleteFolder(path, true);
            }
        } catch (e) {
        }
    }


    function clearStaleLock(path) {
        try {
            if (!fso.FolderExists(path)) {
                return;
            }

            var folder = fso.GetFolder(path);
            var created = new Date(folder.DateCreated);
            var age = new Date().getTime() - created.getTime();

            if (age > STALE_LOCK_MS) {
                fso.DeleteFolder(path, true);
            }
        } catch (e) {
        }
    }


    function deletePendingPackageWithRetry(path) {
        var recoveryPath = recoveryPathFromPending(path);

        /*
          Delete Recovery first. If that fails, keep Primary so the form still
          sees an unsent Pending. If Recovery is deleted but Primary deletion
          fails, the publication receipt remains available on retry.
        */
        if (!deleteFileWithRetry(recoveryPath)) {
            return false;
        }

        return deleteFileWithRetry(path);
    }


    function deleteFileWithRetry(path) {
        var i;

        if (!fso.FileExists(path)) {
            return true;
        }

        for (i = 0; i < 5; i++) {
            try {
                fso.DeleteFile(path, true);
                return true;
            } catch (e) {
                WScript.Sleep(250 + Math.floor(Math.random() * 500));
            }
        }

        return !fso.FileExists(path);
    }


    function showSuccessPopupWhenReady(pendingPath) {
        var readyPath = String(pendingPath || "") + COMPLETION_READY_SUFFIX;
        var startedAt = new Date().getTime();

        while (!fso.FileExists(readyPath) &&
               (new Date().getTime() - startedAt) < COMPLETION_READY_MAX_MS) {
            WScript.Sleep(COMPLETION_READY_POLL_MS);
        }

        showSuccessPopup();

        try {
            if (fso.FileExists(readyPath)) {
                fso.DeleteFile(readyPath, true);
            }
        } catch (e) {
        }

        removeEmptyPendingFolder(pendingPath);
    }


    function showSuccessPopup() {
        try {
            var message =
                "\u30e1\u30f3\u30b7\u30e7\u30f3\u4f9d\u983c\u304c\u5b8c\u4e86\u3057\u307e\u3057\u305f";

            // Timeout 0: keep the completion notice visible until OK is pressed.
            // 64      = MB_ICONINFORMATION
            // 65536   = MB_SETFOREGROUND
            // 262144  = MB_TOPMOST
            // This brings the completion notice to the front even when the main HTA is minimized.
            shell.Popup(
                message,
                0,
                "\u30e1\u30f3\u30b7\u30e7\u30f3\u4f9d\u983c\u30d5\u30a9\u30fc\u30e0 | \u9001\u4fe1\u5b8c\u4e86",
                327744
            );
        } catch (e) {
            // Notification failure must never affect CSV delivery.
        }
    }


    function removeEmptyPendingFolder(path) {
        try {
            var desktop = String(shell.SpecialFolders("Desktop") || "");
            if (!desktop) {
                return;
            }

            var folderPath = fso.GetAbsolutePathName(fso.GetParentFolderName(path));
            var expectedPath = fso.GetAbsolutePathName(fso.BuildPath(desktop, PENDING_FOLDER_NAME));
            if (folderPath.toLowerCase() !== expectedPath.toLowerCase()) {
                return;
            }

            var commandPath = fso.BuildPath(String(fso.GetSpecialFolder(1)), "cmd.exe");
            // Run expands %variables%; reject paths that could be reinterpreted.
            if (/[\x00-\x1f"%]/.test(folderPath) || /[\x00-\x1f"%]/.test(commandPath)) {
                return;
            }

            for (var attempt = 0; attempt < 3; attempt++) {
                if (!fso.FolderExists(folderPath)) {
                    return;
                }
                var folder = fso.GetFolder(folderPath);
                // Do not follow junctions, or remove any files/subdirectories.
                if ((folder.Attributes & 1024) || folder.Files.Count !== 0 || folder.SubFolders.Count !== 0) {
                    return;
                }
                folder = null;

                // No /s: rd must fail if another sender creates a file now.
                // /d disables AutoRun; /v:off keeps ! in quoted paths literal.
                shell.Run('"' + commandPath + '" /d /v:off /c rd "' + folderPath + '"', 0, true);
                if (!fso.FolderExists(folderPath)) {
                    return;
                }
                if (attempt < 2) {
                    WScript.Sleep(200);
                }
            }
        } catch (e) {
            // Cleanup failure must never change delivered CSV or pending data.
        }
    }


    function shouldWaitForFailureReady(pendingPath) {
        try {
            if (!pendingPath || !fso.FileExists(pendingPath)) {
                return false;
            }

            var file = fso.GetFile(pendingPath);
            var createdAt = new Date(file.DateCreated).getTime();
            var age = new Date().getTime() - createdAt;

            return age >= 0 && age <= FAILURE_READY_RECENT_MS;
        } catch (e) {
            return false;
        }
    }


    function showFailurePopupWhenReady(pendingPath, step, error) {
        var readyPath = String(pendingPath || "") + COMPLETION_READY_SUFFIX;
        var startedAt = new Date().getTime();

        // For a just-submitted request, avoid overlapping the acceptance modal.
        // Startup/manual retries are normally old Pending files, so notify immediately.
        if (shouldWaitForFailureReady(pendingPath)) {
            while (!fso.FileExists(readyPath) &&
                   (new Date().getTime() - startedAt) < COMPLETION_READY_MAX_MS) {
                WScript.Sleep(COMPLETION_READY_POLL_MS);
            }
        }

        showFailurePopup(step, error);

        try {
            if (fso.FileExists(readyPath)) {
                fso.DeleteFile(readyPath, true);
            }
        } catch (e) {
        }
    }


    function showFailurePopup(step, error) {
        try {
            var message =
                "\u30e1\u30f3\u30b7\u30e7\u30f3\u4f9d\u983c\u306e\u9001\u4fe1\u3092\u5b8c\u4e86\u3067\u304d\u307e\u305b\u3093\u3067\u3057\u305f\u3002" +
                "\r\n\r\n" +
                "\u672a\u9001\u4fe1\u30c7\u30fc\u30bf\u306f\u4fdd\u5b58\u3055\u308c\u3066\u3044\u307e\u3059\u3002" +
                "\r\n" +
                "\u30d5\u30a9\u30fc\u30e0\u3092\u958b\u304d\u3001\u300c\u672a\u9001\u4fe1\u3092\u518d\u9001\u300d\u304b\u3089\u518d\u9001\u3057\u3066\u304f\u3060\u3055\u3044\u3002";

            // Show the failed step and code, without paths or CSV contents.
            var labels = {
                folder: "\u4fdd\u5b58\u5148\u30d5\u30a9\u30eb\u30c0\u306e\u78ba\u8a8d",
                lock: "\u9000\u907f\u30d5\u30a1\u30a4\u30eb\u306e\u30ed\u30c3\u30af",
                pending: "\u9000\u907fCSV\u306e\u8aad\u307f\u8fbc\u307f",
                name: "\u51fa\u529b\u30d5\u30a1\u30a4\u30eb\u540d\u306e\u6c7a\u5b9a",
                stage: "\u4fdd\u5b58\u5148\u3067\u306e\u4e00\u6642\u30d5\u30a1\u30a4\u30eb\u4f5c\u6210",
                copy: "\u4e00\u6642\u30d5\u30a1\u30a4\u30eb\u3078\u306e\u30b3\u30d4\u30fc",
                verify: "\u30b3\u30d4\u30fc\u5185\u5bb9\u306e\u78ba\u8a8d",
                receipt: "\u518d\u9001\u60c5\u5831\u306e\u8aad\u307f\u8fbc\u307f\u30fb\u4fdd\u5b58",
                rename: "\u4e00\u6642\u30d5\u30a1\u30a4\u30eb\u3092CSV\u540d\u306b\u5909\u66f4",
                cleanup: "\u4fdd\u5b58\u5f8c\u306e\u9000\u907f\u30c7\u30fc\u30bf\u306e\u7247\u4ed8\u3051"
            };
            if (labels[step]) {
                var code = error && error.number ? String(error.number & 65535) : "";
                if (!code && error) {
                    var detail = String(error.message || error.description || "");
                    if (/^(PENDING_|CSV_|OUTPUT_|PUBLISH_|STAGED_|DATE_|HEADER_)[A-Z0-9_]+$/.test(detail)) {
                        code = detail;
                    }
                }
                message += "\r\n\r\n\u505c\u6b62\u7b87\u6240\uff1a" + labels[step];
                if (code) {
                    message += "\uff08\u30a8\u30e9\u30fc\u30b3\u30fc\u30c9\uff1a" + code + "\uff09";
                }
            }

            // 16      = MB_ICONERROR
            // 65536   = MB_SETFOREGROUND
            // 262144  = MB_TOPMOST
            // Timeout 0 keeps the failure notice visible until OK is pressed.
            shell.Popup(
                message,
                0,
                "\u30e1\u30f3\u30b7\u30e7\u30f3\u4f9d\u983c\u30d5\u30a9\u30fc\u30e0 | \u9001\u4fe1\u5931\u6557",
                327696
            );
        } catch (e) {
            // Notification failure must never alter Pending/CSV state.
        }
    }


    function buildOutputStem(requester, sentAt, requestId) {
        var value = String(sentAt || "").replace(/[\uff10-\uff19]/g, function (digit) {
            return String.fromCharCode(digit.charCodeAt(0) - 0xfee0);
        }).replace(/\uff0f/g, "/").replace(/\uff1a/g, ":")
            .replace(/^[\s\u3000\ufeff]+|[\s\u3000\ufeff]+$/g, "");
        var m = /^([0-9]{4})[\/.-]([0-9]{1,2})[\/.-]([0-9]{1,2})[Tt\s\u3000]+([0-9]{1,2}):([0-9]{1,2})(?::[0-9]{1,2}(?:\.[0-9]+)?)?$/.exec(value);
        var stamp = m ? m[1] + ("0" + m[2]).slice(-2) + ("0" + m[3]).slice(-2) +
            ("0" + m[4]).slice(-2) + ("0" + m[5]).slice(-2) : "";

        // A normal send uses the CSV timestamp. For an unrecognized display
        // format, use the time embedded in the original per-request ID.
        if (!isValidOutputMinute(stamp)) {
            var idTime = /^([0-9]{8})-([0-9]{4})[0-9]{5}-.*-R0[1-3]-[0-9A-F]{6}$/i.exec(String(requestId || ""));
            stamp = idTime ? idTime[1] + idTime[2] : "";
            if (!isValidOutputMinute(stamp)) {
                throw new Error("DATE_INVALID");
            }
        }

        var name = String(requester || "").replace(/[\x00-\x1f\\\/:*?"<>|]/g, "_")
            .replace(/^\s+|[\s.]+$/g, "");
        if (!name) {
            name = "\u4f9d\u983c\u8005\u4e0d\u660e";
        }
        if (/^(CON|PRN|AUX|NUL|COM[1-9\u00b9\u00b2\u00b3]|LPT[1-9\u00b9\u00b2\u00b3])\s*\./i.test(name)) {
            name = "_" + name;
        }
        return name + stamp;
    }


    function isValidOutputMinute(stamp) {
        if (stamp.length !== 12) {
            return false;
        }
        var year = Number(stamp.substring(0, 4));
        var month = Number(stamp.substring(4, 6));
        var day = Number(stamp.substring(6, 8));
        return year >= 1000 && month >= 1 && month <= 12 &&
            day >= 1 && day <= new Date(year, month, 0).getDate() &&
            Number(stamp.substring(8, 10)) < 24 && Number(stamp.substring(10, 12)) < 60;
    }


    function readText(path) {
        var stream = fso.OpenTextFile(
            path,
            FOR_READING,
            false,
            TRISTATE_USE_DEFAULT
        );

        try {
            return String(stream.ReadAll() || "");
        } finally {
            stream.Close();
        }
    }


    function normalizeHeaderField(value, index) {
        var actual = String(value || "");
        if (index === 0 && actual.length > 0 && actual.charCodeAt(0) === 0xFEFF) {
            actual = actual.substring(1);
        }
        return actual;
    }


    function headerMatches(row, header) {
        var i;
        if (!row || row.length !== header.length) {
            return false;
        }
        for (i = 0; i < header.length; i++) {
            if (normalizeHeaderField(row[i], i) !== header[i]) {
                return false;
            }
        }
        return true;
    }


    function detectHeader(row) {
        if (headerMatches(row, HEADER_PROXY8)) { return "proxy8"; }
        if (headerMatches(row, HEADER_LATEST)) { return "latest"; }
        if (headerMatches(row, HEADER_V2823)) { return "v2823"; }
        if (headerMatches(row, HEADER_V2822)) { return "v2822"; }
        if (headerMatches(row, HEADER_LEGACY)) { return "legacy"; }
        throw new Error("HEADER_MISMATCH");
    }


    function expectedColumnCount(schema) {
        if (schema === "proxy8") { return HEADER_PROXY8.length; }
        if (schema === "latest") { return HEADER_LATEST.length; }
        if (schema === "v2823") { return HEADER_V2823.length; }
        if (schema === "v2822") { return HEADER_V2822.length; }
        return HEADER_LEGACY.length;
    }


    function sentAtFromRow(row, schema) {
        return String(row[0] || "");
    }


    function baseNameFromRow(row, schema) {
        return String(row[(schema === "proxy8" || schema === "latest" || schema === "v2823") ? 1 : 2] || "");
    }


    function requesterFromRow(row, schema) {
        return String(row[(schema === "proxy8" || schema === "latest" || schema === "v2823") ? 2 : 3] || "");
    }


    function requestNoFromRow(row, schema) {
        return String(row[(schema === "proxy8" || schema === "latest" || schema === "v2823") ? 3 : 4] || "");
    }


    function requestIdFromRow(row, schema) {
        if (schema === "proxy8") { return String(row[20] || ""); }
        if (schema === "latest") { return String(row[18] || ""); }
        if (schema === "v2823") { return String(row[15] || ""); }
        return String(row[1] || "");
    }


    function parseCsv(text) {
        text = String(text == null ? "" : text);

        var rows = [];
        var row = [];
        var field = "";
        var inQuotes = false;
        var i, ch, next;

        for (i = 0; i < text.length; i++) {
            ch = text.charAt(i);

            if (inQuotes) {
                if (ch === '"') {
                    next = (i + 1 < text.length) ? text.charAt(i + 1) : "";

                    if (next === '"') {
                        field += '"';
                        i++;
                    } else {
                        inQuotes = false;
                    }
                } else {
                    field += ch;
                }

                continue;
            }

            if (ch === '"') {
                inQuotes = true;
            } else if (ch === ",") {
                row.push(field);
                field = "";
            } else if (ch === "\r" || ch === "\n") {
                row.push(field);
                field = "";
                rows.push(row);
                row = [];

                if (ch === "\r" &&
                    i + 1 < text.length &&
                    text.charAt(i + 1) === "\n") {
                    i++;
                }
            } else {
                field += ch;
            }
        }

        if (inQuotes) {
            throw new Error("CSV_UNCLOSED_QUOTE");
        }

        if (field !== "" || row.length > 0) {
            row.push(field);
            rows.push(row);
        }

        return rows;
    }


    function isEmptyRecord(row) {
        var i;

        if (!row || row.length === 0) {
            return true;
        }

        for (i = 0; i < row.length; i++) {
            if (String(row[i] || "") !== "") {
                return false;
            }
        }

        return true;
    }


    function sleepRandom() {
        var span = RETRY_MAX_MS - RETRY_MIN_MS + 1;
        var ms = RETRY_MIN_MS + Math.floor(Math.random() * span);
        WScript.Sleep(ms);
    }
})();
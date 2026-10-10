/*
 * 業務量モニター HTML版 v0.8.3 / 平日ログ集計の設定
 * 変更後は保存し、HTMLを開き直してください（UTF-8で保存）。
 * ファイル・親フォルダ・共有先は画面で選択します。列はExcelの列記号で指定します。
 * カンマ、引用符、かっこの削除に注意してください。
 */
var WorkloadConfig = {
    unread: {
        startRow: 3,                    // 1・2行目を除外
        charset: "auto",               // auto / utf-8 / shift_jis / unicode
        columns: {
            organization: "C",
            caName: "E",
            unreadBase: "H",
            unreadSubtract: ["M", "N", "O"],
            unsorted: "R"
        }
    },
    flag: {
        startRow: 3,
        charset: "auto",
        columns: {
            organization: "C",
            caName: "E",
            flag: "F"
        }
    },
    assignment: {
        sheetName: "CA振分表",          // この名前のシートを読みます
        startRow: 3,
        columns: {
            organization: "C",
            caName: "E",
            calib: "F",
            team: "K",
            staff: "L",
            canvas: "M"
        }
    },
    capacity: {
        hourlyRate: 16,
        excludedStaffNames: ["担当なし", "札幌担当", "新潟担当", "首都圏担当"],
        // チーム欄に下記の区分がある担当者も、対応余力の人数から除外。
        excludedTeams: ["担当なし", "札幌", "札幌担当", "新潟", "新潟担当", "首都圏", "首都圏担当"]
    },
    merge: {
        unassignedStaffName: "担当なし",
        unassignedTeam: "担当なし",
        // 同じ組織・CA名がCSVに複数行ある場合：sumで合算、errorで停止。
        duplicateLogRows: "sum"
    }
};

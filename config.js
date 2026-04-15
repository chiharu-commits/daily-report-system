// 日報システムの設定データ
const CONFIG = {
  // 環境設定（テスト環境: 'test', 本番環境: 'production'）
  "environment": "production",

  // Google Apps Script URL
  "scriptUrls": {
    "test": "https://script.google.com/macros/s/AKfycbzlatpQDr1p-Pkj8zB50lFYiurXdFT6R64aaDyhG3I2JEppxkHxgMDP7yWVgRloA2IQ/exec",
    "production": "https://script.google.com/macros/s/AKfycbyf8fJx5znNDEFfXxpaZE75HiWs8QSM3BKuZuPu_iJ3UWNwtjqrMpSzxAHg2sAXBLfSyA/exec"
  },

  "names": [
    "田代恭子",
    "佐々木由希江",
    "山本千春"
  ],
  "categories": [
    "NEWTON対応",
    "運用保守支援",
    "会議",
    "AI開発",
    "休み"
  ],
  "taskNames": [
    "NEWTON受入テスト",
    "NEWTON本番移行",
    "後追いKEY採番",
    "レビュー管理作業",
    "MDM月次報告会",
    "打合せ/レビュー",
    "進捗会議",
    "開発作業",
    "休暇"
  ]
};

// 現在の環境のGoogle Apps Script URLを取得
function getScriptUrl() {
  return CONFIG.scriptUrls[CONFIG.environment];
}

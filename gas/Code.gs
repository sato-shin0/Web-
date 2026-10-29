/**
 * 学級旗コンテスト投票アプリ：学級旗画像の共有保存用 Google Apps Script
 *
 * 管理画面で変更した学級旗画像を Google ドライブに保存し、
 * すべての端末から同じ画像を表示できるようにします。
 *
 * 【設定手順】
 * 1. https://script.google.com/ で「新しいプロジェクト」を作成
 * 2. このファイルの内容をすべて貼り付けて保存
 * 3. 右上「デプロイ」→「新しいデプロイ」→ 種類「ウェブアプリ」を選択
 *      - 次のユーザーとして実行：自分
 *      - アクセスできるユーザー：全員
 *    で「デプロイ」し、ドライブへのアクセスを許可する
 * 4. 表示された「ウェブアプリのURL（…/exec）」をコピーし、
 *    投票アプリ.html の SETTINGS_API_URL に貼り付ける
 *
 * ※ このスクリプトを修正した場合は「デプロイを管理」から新しいバージョンを発行してください。
 * ※ ADMIN_PASSWORD は投票アプリ.html の管理者パスワードと同じ値にしてください。
 */

const ADMIN_PASSWORD = 'admin123';
const FOLDER_NAME = '学級旗コンテスト画像';
const PROP_KEY = 'flagImages'; // { flagId: driveFileId }
const FLAG_IDS = ['1-1', '1-2', '1-3', '2-1', '2-2', '3-1', '3-2', '3-3'];

// 全端末向け：保存済みの学級旗画像を返す
function doGet() {
  const map = getImageMap_();
  const images = {};
  Object.keys(map).forEach(id => {
    try {
      const blob = DriveApp.getFileById(map[id]).getBlob();
      images[id] = 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
    } catch (err) {
      // ファイルが手動で削除された場合などは初期画像を表示させる
    }
  });
  return json_({ ok: true, images: images });
}

// 管理画面向け：画像の保存 / 初期化
function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const body = JSON.parse(e.postData.contents);
    if (body.password !== ADMIN_PASSWORD) {
      return json_({ ok: false, error: 'パスワードが正しくありません' });
    }

    let map = getImageMap_();

    if (body.action === 'save') {
      const folder = getFolder_();
      Object.keys(body.images || {}).forEach(id => {
        if (FLAG_IDS.indexOf(id) === -1) return;
        const m = /^data:(image\/[\w+.-]+);base64,(.+)$/.exec(body.images[id]);
        if (!m) return;
        const blob = Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], id + '.jpg');
        const file = folder.createFile(blob);
        if (map[id]) trashQuietly_(map[id]);
        map[id] = file.getId();
      });
    } else if (body.action === 'reset') {
      Object.keys(map).forEach(id => trashQuietly_(map[id]));
      map = {};
    } else {
      return json_({ ok: false, error: '不明な操作です' });
    }

    PropertiesService.getScriptProperties().setProperty(PROP_KEY, JSON.stringify(map));
    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

function getImageMap_() {
  return JSON.parse(PropertiesService.getScriptProperties().getProperty(PROP_KEY) || '{}');
}

function getFolder_() {
  const folders = DriveApp.getFoldersByName(FOLDER_NAME);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(FOLDER_NAME);
}

function trashQuietly_(fileId) {
  try {
    DriveApp.getFileById(fileId).setTrashed(true);
  } catch (err) {
    // 既に削除済み
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

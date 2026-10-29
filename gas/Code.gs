/**
 * 学級旗コンテスト投票アプリ：共有データ保存用 Google Apps Script
 *
 * 管理画面で変更した学級旗画像（Google ドライブ）と投票データ（Google スプレッドシート）を
 * 一か所に保存し、すべての端末で同じ画像・同じ集計を使えるようにします。
 * 生徒の二重投票チェックもここで行うため、別の端末からの再投票も防げます。
 *
 * 【設定手順】
 * 1. https://script.google.com/ で「新しいプロジェクト」を作成
 * 2. このファイルの内容をすべて貼り付け、下の ADMIN_PASSWORD を独自のパスワードに変更して保存
 * 3. 右上「デプロイ」→「新しいデプロイ」→ 種類「ウェブアプリ」を選択
 *      - 次のユーザーとして実行：自分
 *      - アクセスできるユーザー：全員
 *    で「デプロイ」し、ドライブ・スプレッドシートへのアクセスを許可する
 * 4. 表示された「ウェブアプリのURL（…/exec）」をコピーし、
 *    投票アプリ.html の API_URL に貼り付ける
 *
 * ※ このスクリプトを修正した場合（パスワード変更を含む）は、
 *    「デプロイ」→「デプロイを管理」→ 編集 → バージョン「新バージョン」で再デプロイしてください。
 * ※ 画像と投票データは、マイドライブの「学級旗コンテスト画像」フォルダに保存されます。
 */

// 管理画面のパスワード（必ず変更してください。HTML 側には書かれないため外部から見えません）
const ADMIN_PASSWORD = 'change-me';

const FOLDER_NAME = '学級旗コンテスト画像';
const SPREADSHEET_NAME = '学級旗コンテスト投票データ';
const PROP_KEY = 'flagImages';               // { flagId: driveFileId }
const PROP_SPREADSHEET_ID = 'voteSpreadsheetId';
const FLAG_IDS = ['1-1', '1-2', '1-3', '2-1', '2-2', '3-1', '3-2', '3-3'];
const ROLE_LABELS = { student: '生徒', parent: '保護者・一般' };
const VOTE_HEADERS = ['投票ID', '区分', '生徒(学年-組-番号)', '投票先', '投票日時'];

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

function doPost(e) {
  let body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'bad_request' });
  }

  const publicActions = {
    checkStudent: checkStudent_,
    vote: vote_
  };
  const adminActions = {
    login: () => ({ ok: true }),
    getVotes: getVotes_,
    resetVotes: resetVotes_,
    save: saveImages_,
    reset: resetImages_
  };

  if (publicActions[body.action]) {
    return json_(withLock_(() => publicActions[body.action](body)));
  }
  if (adminActions[body.action]) {
    if (ADMIN_PASSWORD === 'change-me') {
      return json_({ ok: false, error: 'Code.gs の ADMIN_PASSWORD を変更して再デプロイしてください' });
    }
    if (body.password !== ADMIN_PASSWORD) {
      return json_({ ok: false, error: 'unauthorized' });
    }
    return json_(withLock_(() => adminActions[body.action](body)));
  }
  return json_({ ok: false, error: 'unknown_action' });
}

// ---------- 投票 ----------

function checkStudent_(body) {
  if (!isValidStudentId_(body.studentId)) return { ok: false, error: 'invalid_student' };
  return { ok: true, voted: getVotedStudentIds_().indexOf(body.studentId) !== -1 };
}

function vote_(body) {
  if (!ROLE_LABELS[body.role]) return { ok: false, error: 'invalid_role' };
  if (FLAG_IDS.indexOf(body.flagId) === -1) return { ok: false, error: 'invalid_flag' };

  let studentId = '';
  if (body.role === 'student') {
    if (!isValidStudentId_(body.studentId)) return { ok: false, error: 'invalid_student' };
    if (getVotedStudentIds_().indexOf(body.studentId) !== -1) return { ok: false, error: 'already_voted' };
    studentId = body.studentId;
  }

  const record = {
    id: Date.now().toString() + Math.floor(Math.random() * 1000),
    role: body.role,
    flagId: body.flagId,
    studentId: studentId || null,
    timestamp: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss')
  };
  getVoteSheet_().appendRow([record.id, ROLE_LABELS[record.role], studentId, record.flagId, record.timestamp]);
  return { ok: true, record: record };
}

function getVotes_() {
  const sheet = getVoteSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: true, votes: [] };

  const roleByLabel = {};
  Object.keys(ROLE_LABELS).forEach(key => { roleByLabel[ROLE_LABELS[key]] = key; });

  const votes = sheet.getRange(2, 1, lastRow - 1, VOTE_HEADERS.length).getDisplayValues().map(row => ({
    id: row[0],
    role: roleByLabel[row[1]] || row[1],
    studentId: row[2] || null,
    flagId: row[3],
    timestamp: row[4]
  }));
  return { ok: true, votes: votes };
}

function resetVotes_() {
  const sheet = getVoteSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) sheet.deleteRows(2, lastRow - 1);
  return { ok: true };
}

function getVotedStudentIds_() {
  const sheet = getVoteSheet_();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 3, lastRow - 1, 1).getDisplayValues().map(row => row[0]).filter(String);
}

function isValidStudentId_(studentId) {
  return typeof studentId === 'string' && /^[1-3]-\d-\d{1,2}$/.test(studentId);
}

function getVoteSheet_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty(PROP_SPREADSHEET_ID);
  if (id) {
    try {
      return SpreadsheetApp.openById(id).getSheets()[0];
    } catch (err) {
      // スプレッドシートが削除された場合は作り直す
    }
  }

  const ss = SpreadsheetApp.create(SPREADSHEET_NAME);
  DriveApp.getFileById(ss.getId()).moveTo(getFolder_());
  const sheet = ss.getSheets()[0];
  sheet.setName('投票');
  // 「1-1-5」などが日付に変換されないよう、すべて書式なしテキストにする
  sheet.getRange(1, 1, sheet.getMaxRows(), VOTE_HEADERS.length).setNumberFormat('@');
  sheet.getRange(1, 1, 1, VOTE_HEADERS.length).setValues([VOTE_HEADERS]).setFontWeight('bold');
  sheet.setFrozenRows(1);
  props.setProperty(PROP_SPREADSHEET_ID, ss.getId());
  return sheet;
}

// ---------- 学級旗画像 ----------

function saveImages_(body) {
  const map = getImageMap_();
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
  PropertiesService.getScriptProperties().setProperty(PROP_KEY, JSON.stringify(map));
  return { ok: true };
}

function resetImages_() {
  const map = getImageMap_();
  Object.keys(map).forEach(id => trashQuietly_(map[id]));
  PropertiesService.getScriptProperties().setProperty(PROP_KEY, '{}');
  return { ok: true };
}

function getImageMap_() {
  return JSON.parse(PropertiesService.getScriptProperties().getProperty(PROP_KEY) || '{}');
}

// ---------- 共通 ----------

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return fn();
  } finally {
    lock.releaseLock();
  }
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

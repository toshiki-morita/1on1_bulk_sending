/** 
 * 1) スプレッドシートを用意（ExcelはDriveでGSに変換） 
 * 2) 下のCONFIGを書き換え
 * 3) メニュー > 実行: createDraftsFromAttackList()
 */

const CONFIG = {
  // 元データ（要件定義のID）
  SHEET_ID: '1o8lo5HbvdpcH2YtEjZ7jXj5o0jcPA7-2QuYsVG3iRfc',
  SRC_SHEET_NAME: 'アタックリスト',
  CONTACTS_SHEET_NAME: 'Contacts',

  // ログタブ
  SEND_LOG_SHEET_NAME: '送信ログ',   // 成功/スキップ/エラーを記録
  LOG_SHEET_NAME: '未送信リスト',    // 宛先不足など（既存運用）

  // メール設定
  DEADLINE_TEXT: '8月22日',
  SENDER_NAME: 'ユビ電　営業サポート',
  ATTACH_FILE_ID: '<<AI音声付きチラシPDFのDriveファイルID>>',
  ATTACH_URL: 'https://wecharge-my.sharepoint.com/:b:/g/personal/toshiki_morita_ubiden_com/EX_wRly2pJNLtaRLwyJU8vcB8dGA3U2400dYWmAgam9yug?e=Mx0pZN',
  DRY_RUN: false, // true: 下書き作成 / false: 送信
  SUBJECT: '[テスト！返信してください！] 各マンション理事会様へのご案内状況の確認',
  MAX_ITEMS_PER_MAIL: 10, // 10件超は分割
  APPLY_SENT_FLAG: true, // 送信後に「送信済」列へ日時を記録
  SENT_FLAG_COL_CANDIDATES: ['送信済', '送信済フラグ'],
  FROM_ALIAS: 'sales@ubiden.com', // Gmailの「別名として送信」に設定済みなら使用
  REPLY_TO: 'sales@ubiden.com', // 返信先（Fromエイリアス未設定時の代替）
  LABEL_NAME: '1on1/過去案件リマインド', // 空ならラベル付与なし

  // 担当者メール列の候補名（保険）
  EMAIL_COL_CANDIDATES: ['メール', 'メールアドレス', 'フロント担当者メール', 'Email'],

  // 抽出条件
  FILTER: {
    // 「対応した？」チェックボックスが TRUE のものは除外し、FALSE/空のみ対象
    respondedCol: '対応した？',
    proposalCol: '提案可否',
    proposalEquals: '提案可',
    involvementCol: '次回理事会の関与形式',
    involvementMustIncludes: 'チラシ案内'
  },

  // Contacts列
  CONTACT_KEY_SEP: '｜',
  CONTACTS_COLS: {
    key: 'contact_key',
    to: 'フロント担当者メール',
    cc1: 'ユビ電セールス担当メール（CC）1',
    cc2: 'ユビ電セールス担当メール（CC）2',
    nameSales: '自社営業担当名'
  }
};

function createDraftsFromAttackList() {
  // 1) データ取得
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const contacts = loadContactsMap_(ss); // ★Contactsマップ
  const sheet = ss.getSheetByName(CONFIG.SRC_SHEET_NAME);
  if (!sheet) throw new Error('シートが見つかりません: ' + CONFIG.SRC_SHEET_NAME);
  const rows = readAsObjects_(sheet);
  const filtered = filterTargets_(rows);

  // 2) グルーピング・ログシート用意
  const grouped = groupByContactKey_(filtered);
  const sendLogSheet = ensureLogSheet_(ss, CONFIG.SEND_LOG_SHEET_NAME);
  const unsentLogSheet = ensureLogSheet_(ss, CONFIG.LOG_SHEET_NAME);
  ensureUnsentHeader_(unsentLogSheet);

  // 3) 添付/リンク準備
  let attachBlob = null;
  let attachLink = '';
  try {
    const fileIdOrUrl = CONFIG.ATTACH_FILE_ID;
    if (fileIdOrUrl) {
      if (isHttpUrl_(fileIdOrUrl)) {
        attachLink = fileIdOrUrl;
      } else {
        attachBlob = DriveApp.getFileById(fileIdOrUrl).getBlob();
      }
    }
  } catch (e) {
    // 取得失敗は致命ではないため通知のみ（リンクがあれば続行）
    SpreadsheetApp.getActiveSpreadsheet().toast('添付取得に失敗: ' + e.message, '過去案件メール', 7);
  }
  if (CONFIG.ATTACH_URL) attachLink = CONFIG.ATTACH_URL;

  // 3.5) ラベル準備（必要時）
  let label = null;
  try {
    if (CONFIG.LABEL_NAME) {
      label = ensureGmailLabel_(CONFIG.LABEL_NAME);
    }
  } catch (e) {
    SpreadsheetApp.getActiveSpreadsheet().toast('ラベル準備に失敗: ' + e.message, '過去案件メール', 5);
  }

  // 4) 送信/下書き作成
  let created = 0, skipped = 0, errors = 0;
  Object.values(grouped).forEach(items => {
    const sample = items[0];
    const key = makeContactKey_(sample);
    const c = contacts.get(key);

    if (!c || !isLikelyEmail_(c.to)) {
      logUnsent_(unsentLogSheet, '宛先なし（Contacts未登録）', sample, sheet.getName());
      logSend_(sendLogSheet, { company: String(sample['会社名'] || ''), branch: String(sample['支店・部署'] || ''), frontName: String(sample['フロント担当者名'] || ''), to: '', cc: '', count: items.length, mode: CONFIG.DRY_RUN ? 'draft' : 'send', status: 'skip', reason: '宛先なし' });
      skipped++;
      return;
    }

    const chunks = chunkArray_(items, CONFIG.MAX_ITEMS_PER_MAIL);
    chunks.forEach((chunk, idx) => {
      const companyBranch = `${sample['会社名'] || ''} ${sample['支店・部署'] || ''}`.trim();
      const userName = sample['フロント担当者名'] || 'ご担当者';
      const htmlBody = buildMailHtml_(companyBranch, userName, chunk, attachLink, !!attachBlob);
      const plain = stripHtml_(htmlBody);

      // CCは最大2名まで（Contactsの2列）。妥当なメールのみカンマ連結
      const ccList = [c.cc1, c.cc2].filter(e => isLikelyEmail_(e)).join(',');
      const options = { htmlBody, name: CONFIG.SENDER_NAME };
      if (attachBlob) options.attachments = [attachBlob];
      if (ccList) options.cc = ccList;
      // Fromエイリアス（sales@ubiden.com）が使える場合はfromを指定。なければreplyToへフォールバック
      try {
        const aliases = GmailApp.getAliases && GmailApp.getAliases();
        if (CONFIG.FROM_ALIAS && Array.isArray(aliases) && aliases.indexOf(CONFIG.FROM_ALIAS) !== -1) {
          options.from = CONFIG.FROM_ALIAS;
        }
      } catch (_) {}
      if (CONFIG.REPLY_TO) options.replyTo = CONFIG.REPLY_TO;

      let subject = CONFIG.SUBJECT;
      if (chunks.length > 1) subject += ` (${idx + 1}/${chunks.length})`;

      try {
        // まずはドラフトを作成
        const draft = GmailApp.createDraft(c.to, subject, plain, options);
        // ラベル付与（スレッド単位）＋スレッドID保持
        var draftThreadId = null;
        try {
          if (label && draft && draft.getMessage) {
            const thread = draft.getMessage().getThread();
            if (thread && thread.addLabel) thread.addLabel(label);
            if (thread && thread.getId) draftThreadId = thread.getId();
          }
        } catch (e2) {
          // ラベル付与失敗は致命ではないため通知のみ
          SpreadsheetApp.getActiveSpreadsheet().toast('ラベル付与に失敗: ' + e2.message, '過去案件メール', 5);
        }
        // 本番時は送信、DRY_RUN時はドラフト保持
        if (!CONFIG.DRY_RUN) {
          draft.send();
          // 送信直後、保持したスレッドIDでラベル再付与（より堅牢）
          try {
            if (label && draftThreadId) {
              var t = GmailApp.getThreadById(draftThreadId);
              if (t && t.addLabel) t.addLabel(label);
            }
          } catch (e2b) {
            SpreadsheetApp.getActiveSpreadsheet().toast('送信後ラベル付与(スレッドID)失敗: ' + e2b.message, '過去案件メール', 4);
          }
          // フォールバック: 直近の送信済みスレッドを検索してラベル付与（遅延＋リトライ）
          try {
            if (label) {
              Utilities.sleep(1000); // 反映待ち（短時間）
              var subj = subject.replace(/"/g, '\\"');
              var fromQ = (options && options.from) ? (' from:' + options.from) : '';
              var qBase = 'in:sent' + fromQ + ' to:' + c.to + ' subject:"' + subj + '" newer_than:2d';
              var threads = [];
              for (var i = 0; i < 3; i++) {
                threads = GmailApp.search(qBase, 0, 1);
                if (threads && threads.length > 0) {
                  threads[0].addLabel(label);
                  break;
                }
                Utilities.sleep(1000);
              }
              if (!threads || threads.length === 0) {
                SpreadsheetApp.getActiveSpreadsheet().toast('送信後ラベル付与(検索)できず: ' + qBase, '過去案件メール', 5);
              }
            }
          } catch (e3) {
            SpreadsheetApp.getActiveSpreadsheet().toast('送信後ラベル付与(検索)失敗: ' + e3.message, '過去案件メール', 4);
          }
        }

        logSend_(sendLogSheet, { company: String(sample['会社名'] || ''), branch: String(sample['支店・部署'] || ''), frontName: String(userName || ''), to: String(c.to || ''), cc: String(ccList || ''), count: chunk.length, mode: CONFIG.DRY_RUN ? 'draft' : 'send', status: 'ok', reason: '' });
        created++;

        if (CONFIG.APPLY_SENT_FLAG && !CONFIG.DRY_RUN) applySentFlagIfNeeded_(sheet, chunk);
      } catch (e) {
        logSend_(sendLogSheet, { company: String(sample['会社名'] || ''), branch: String(sample['支店・部署'] || ''), frontName: String(userName || ''), to: String(c.to || ''), cc: String(ccList || ''), count: chunk.length, mode: CONFIG.DRY_RUN ? 'draft' : 'send', status: 'error', reason: e.message });
        errors++;
      }
    });
  });

  // 未送・エラーがいずれも0件なら、未送信リストをクリア（ヘッダのみ残す）
  if (skipped === 0 && errors === 0) clearUnsentSheet_(unsentLogSheet);

  SpreadsheetApp.getActiveSpreadsheet().toast(`下書き/送信:${created} / スキップ:${skipped} / エラー:${errors}`, '過去案件メール', 7);
}

/**
 * Contactsタブを読み込み、contact_key→{to,cc1,cc2}のMapを返す
 * @param {GoogleAppsScript.Spreadsheet.Spreadsheet} ss
 * @return {Map<string, {to?: string, cc1?: string, cc2?: string}>}
 */
function loadContactsMap_(ss) {
  const sh = ss.getSheetByName(CONFIG.CONTACTS_SHEET_NAME);
  if (!sh) throw new Error('Contactsシートがありません: ' + CONFIG.CONTACTS_SHEET_NAME);
  const rows = readAsObjects_(sh);
  const m = new Map();
  rows.forEach(r => {
    const key = String(r[CONFIG.CONTACTS_COLS.key] ?? '').trim();
    if (!key) return;
    const to = String(r[CONFIG.CONTACTS_COLS.to] ?? '').trim();
    const cc1 = String(r[CONFIG.CONTACTS_COLS.cc1] ?? '').trim();
    const cc2 = String(r[CONFIG.CONTACTS_COLS.cc2] ?? '').trim();
    if (to) m.set(key, { to, cc1, cc2 });
  });
  return m;
}

/**
 * 会社名｜支店・部署｜フロント担当者名 から contact_key を生成
 * @param {Object} r
 */
function makeContactKey_(r) {
  return [
    String(r['会社名'] ?? ''),
    String(r['支店・部署'] ?? ''),
    String(r['フロント担当者名'] ?? '')
  ].join(CONFIG.CONTACT_KEY_SEP);
}

/**
 * 同一contact_keyでグルーピング
 * @param {Object[]} rows
 * @return {Object<string, Object[]>}
 */
function groupByContactKey_(rows) {
  const map = {};
  rows.forEach(r => {
    const key = makeContactKey_(r);
    if (!map[key]) map[key] = [];
    map[key].push(r);
  });
  return map;
}

/**
 * シートをヘッダ行基準でオブジェクト配列化。__rowに元の行番号を付与。
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 * @return {Object[]}
 */
function readAsObjects_(sheet) {
  const range = sheet.getDataRange();
  const values = range.getValues();
  const headers = values.shift();
  return values
    .filter(r => r.some(v => v !== '' && v !== null))
    .map((row, i) => {
      const obj = Object.fromEntries(headers.map((h, colIdx) => [String(h).trim(), row[colIdx]]));
      obj.__row = i + 2; // データ開始は2行目
      return obj;
    });
}

/**
 * 提案可・チラシ案内・日付上限などでフィルタ。任意で未送信のみ。
 * @param {Object[]} rows
 * @return {Object[]}
 */
function filterTargets_(rows) {
  const f = CONFIG.FILTER;
  let out = rows.slice();

  // 「対応した？」がTRUEのものを除外（チェック未のみ対象）
  if (f.respondedCol) {
    out = out.filter(r => {
      const v = r[f.respondedCol];
      const s = String(v).toLowerCase();
      const isTrue = v === true || s === 'true' || s === '✔' || s === '✓';
      return !isTrue; // FALSE/空は対象
    });
  }

  // 提案可否
  if (f.proposalCol && f.proposalEquals) {
    out = out.filter(r => String(r[f.proposalCol] ?? '').includes(f.proposalEquals));
  }
  // 関与形式（チラシ案内）
  if (f.involvementCol && f.involvementMustIncludes) {
    out = out.filter(r => String(r[f.involvementCol] ?? '').includes(f.involvementMustIncludes));
  }
  return out;
}

function detectEmailCol_(rows) {
  const cands = CONFIG.EMAIL_COL_CANDIDATES;
  for (const c of cands) {
    if (rows.some(r => r[c])) return c;
  }
  // 見つからない場合は空の列名を返し、宛先なしとしてログ
  return '';
}

function groupByContact_(rows, emailCol) {
  // 会社名＋支店・部署＋担当者名＋メール で1通に束ねる
  const map = {};
  rows.forEach(r => {
    const key = [
      String(r['会社名'] ?? ''),
      String(r['支店・部署'] ?? ''),
      String(r['フロント担当者名'] ?? ''),
      emailCol ? String(r[emailCol] ?? '') : ''
    ].join('||');
    (map[key] = map[key] || []).push(r);
  });
  // 案件名で安定ソート
  Object.values(map).forEach(arr => arr.sort((a, b) => String(a['マンション名']).localeCompare(String(b['マンション名']))));
  return map;
}

/**
 * メール本文（HTML）を生成。必要に応じて添付リンクと文言を切替。
 * @param {string} companyBranch 会社名+支店/部署
 * @param {string} userName フロント担当者名
 * @param {Object[]} items 対象案件配列
 * @param {string} attachLink OneDrive等の共有URL（任意）
 * @param {boolean} hasAttach Drive添付が付く場合true
 * @return {string} HTML本文
 */
function buildMailHtml_(companyBranch, userName, items, attachLink, hasAttach) {
  const deadline = CONFIG.DEADLINE_TEXT;

  const rowsHtml = items.map(r => {
    const name = escapeHtml_(String(r['マンション名'] ?? ''));
    const dRaw = r['次の理事会日/日付不明は1日で仮設定'];
    const dStr = formatDateJp_(dRaw);
    return `<tr><td style="padding:6px 10px;border:1px solid #ddd;">${name}</td><td style="padding:6px 10px;border:1px solid #ddd;">${dStr}</td></tr>`;
  }).join('');

  const table = `
  <table style="border-collapse:collapse;border:1px solid #ddd;">
    <thead>
      <tr>
        <th style="padding:6px 10px;border:1px solid #ddd;background:#f6f7f9;text-align:left;">マンション名</th>
        <th style="padding:6px 10px;border:1px solid #ddd;background:#f6f7f9;text-align:left;">当初予定理事会日</th>
      </tr>
    </thead>
    <tbody>${rowsHtml}</tbody>
  </table>`;

  // 本文（要件定義の文面ベース。差出人名と「AI音声付き」に統一）
  const supplement = hasAttach
    ? '<div>■補足</div>'
    : '<div>最新版のAI音声付チラシ（PDF）を添付しております。次回提案時にぜひご活用ください。</div>';
  const linkRow = attachLink
    ? `<div>ダウンロードはこちら: <a href="${escapeHtml_(attachLink)}">${escapeHtml_(attachLink)}</a></div>`
    : '';
  const body = `
<div style="font-size:14px;line-height:1.9;color:#111;">
<div>${escapeHtml_(companyBranch)}</div>
<div>${escapeHtml_(userName)} さま</div>
<br>
<div>いつも大変お世話になっております。${escapeHtml_(CONFIG.SENDER_NAME)}森田です。</div>
<br>
<div>以前、弊社各営業担当者が個別にヒアリングを実施し、各マンション理事会で補助金を活用したEV充電設備導入のご案内を進めていただくようお願いしておりました。</div>
<div>まずはご多忙の中、理事会でのご提案・ご案内をいただき、誠にありがとうございます。</div>
<br>
<div>今回、過去に「ご提案可能」と伺っていた物件のうち、ご案内状況の確認が取れていない案件がいくつかありましたので、大変恐縮ですが、下記の通り状況をご教示いただければ幸いです。</div>
<br>
<div><b>■確認したい内容</b></div>
<ol style="margin:6px 0 14px 20px;">
  <li>提案済みかどうか</li>
  <li>提案済みの場合、その結果（検討可能・保留・見送り など）</li>
  <li>未提案の場合、次回理事会での提案可否
    <ul style="margin:6px 0 6px 18px;">
      <li>提案可能な場合は日程</li>
      <li>難しい場合は理由</li>
    </ul>
  </li>
</ol>
<div><b>■対象物件一覧</b></div>
${table}
<br>
<div>一部、ヒヤリング時に日程が未定だったものは月初1日で仮設定の場合があります。ご容赦ください。</div>
<div>また、ご担当者さまに変更があった場合は、お手数ですが新たなご担当者さまの<b>お名前とメールアドレス</b>も併せてお知らせください。</div>
<br>
<div>お忙しいところ恐れ入りますが、${escapeHtml_(deadline)}までにご返信をお願いいたします。</div>
<div><b>■補足</b></div>
${supplement}
${linkRow}
<div>何卒よろしくお願い申し上げます。</div>
<br>
${buildSignatureHtml_()}
</div>
  `;
  return body.trim();
}

/**
 * 署名（HTML）を生成して返す。
 * @return {string}
 */
function buildSignatureHtml_() {
  return `
<div style="margin-top:16px;font-size:13px;line-height:1.7;color:#333;">
<div>--------------------------------------------------------------------</div>
<div>森田　稔己 / Toshiki Morita</div>
<div>ユビ電株式会社 </div>
<div>ビジネスストラテジー／カスタマーサクセス</div>
<div>〒108-0073　東京都港区三田一丁目1番14号　Bizflex麻布十番4階</div>
<div>TEL 080-7439-7098 </div>
<div>名刺：<a href="https://8card.net/virtual_cards/1jlunfBTRHBP85DHWGqskA">https://8card.net/virtual_cards/1jlunfBTRHBP85DHWGqskA</a></div>
<div>HP： <a href="https://www.ubiden.com">https://www.ubiden.com</a></div>
<div>└───────────────┘</div>
<div>--------------------------------------------------------------------</div>
</div>
  `.trim();
}

/**
 * 最低限のプレーン文生成（GmailApp要件で本文は必須）
 * @param {string} html
 * @return {string}
 */
function stripHtml_(html) {
  return html.replace(/<br\s*\/?>/gi, '\n')
             .replace(/<[^>]+>/g, '')
             .replace(/&nbsp;/g, ' ')
             .replace(/&amp;/g, '&')
             .replace(/&lt;/g, '<')
             .replace(/&gt;/g, '>')
             .trim();
}

function escapeHtml_(s) {
  return s.replace(/&/g, '&amp;')
          .replace(/</g, '&lt;')
          .replace(/>/g, '&gt;');
}

function formatDateJp_(v) {
  if (!v) return '';
  let d;
  if (v instanceof Date) d = v;
  else {
    const t = new Date(v);
    if (isNaN(+t)) return String(v);
    d = t;
  }
  const y = d.getFullYear();
  const m = ('0' + (d.getMonth() + 1)).slice(-2);
  const dd = ('0' + d.getDate()).slice(-2);
  return `${y}/${m}/${dd}`;
}

function ensureLogSheet_(ss, name) {
  return ss.getSheetByName(name) ?? ss.insertSheet(name);
}

function logUnsent_(sheet, reason, row, fileName) {
  const now = new Date();
  sheet.appendRow([
    Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss.SSS'),
    reason,
    fileName,
    row && row.__row ? row.__row : '',
    row ? (row['マンション名'] ?? '') : '',
    row ? (row['フロント担当者名'] ?? '') : '',
  ]);
}

/**
 * 未送信リストのヘッダを保証
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function ensureUnsentHeader_(sheet) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['実行日時','エラー種別','ファイル名','行番号','マンション名','フロント担当者名']);
  }
}

/**
 * 未送信が0件のとき、前回のデータをクリアしてヘッダのみ残す
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet
 */
function clearUnsentSheet_(sheet) {
  sheet.clearContents();
  sheet.appendRow(['実行日時','エラー種別','ファイル名','行番号','マンション名','フロント担当者名']);
}

/**
 * 送信ログを追記
 */
function logSend_(sheet, rec) {
  const now = new Date();
  const row = [
    Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss.SSS'),
    rec.mode || '',
    rec.status || '',
    rec.to || '',
    rec.cc || '',
    rec.count || 0,
    rec.company || '',
    rec.branch || '',
    rec.frontName || '',
    rec.reason || ''
  ];
  sheet.appendRow(row);
}

/**
 * 指定名のGmailラベルを取得（なければ作成）
 * @param {string} name ラベル名（例: "1on1/過去案件リマインド"）
 * @return {GoogleAppsScript.Gmail.GmailLabel} ラベルオブジェクト
 */
function ensureGmailLabel_(name) {
  if (!name) return null;
  var label = GmailApp.getUserLabelByName(name);
  if (!label) {
    label = GmailApp.createLabel(name);
  }
  return label;
}

/**
 * 送信済み列に送信日時を記録（chunk内の行を対象）
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet アタックリストシート
 * @param {Object[]} items 送信対象オブジェクト配列（__rowを含む）
 */
function applySentFlagIfNeeded_(sheet, items) {
  try {
    if (!Array.isArray(items) || items.length === 0) return;
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim());
    const candidates = Array.isArray(CONFIG.SENT_FLAG_COL_CANDIDATES) && CONFIG.SENT_FLAG_COL_CANDIDATES.length > 0
      ? CONFIG.SENT_FLAG_COL_CANDIDATES
      : ['送信済'];
    let colIndex = -1;
    for (const name of candidates) {
      const idx = headers.indexOf(name);
      if (idx !== -1) { colIndex = idx + 1; break; }
    }
    if (colIndex === -1) throw new Error('「送信済」列が見つかりません');

    const now = new Date();
    // 行ごとに日時を記録（chunkは最大10件のため逐次更新で十分）
    items.forEach(it => {
      if (it && it.__row) sheet.getRange(it.__row, colIndex).setValue(now);
    });
  } catch (e) {
    SpreadsheetApp.getActiveSpreadsheet().toast('送信済列更新失敗: ' + e.message, '過去案件メール', 5);
  }
}

/** メールアドレスの簡易検証 */
function isLikelyEmail_(s) {
  if (!s) return false;
  const re = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return re.test(String(s).trim());
}

/**
 * 与えられた文字列が http(s) URL かを簡易判定
 * @param {string} s
 * @return {boolean}
 */
function isHttpUrl_(s) {
  if (!s || typeof s !== 'string') return false;
  return /^https?:\/\//i.test(s.trim());
}

/** 配列を指定サイズで分割 */
function chunkArray_(arr, size) {
  if (!Array.isArray(arr) || size <= 0) return [arr];
  const out = [];
  for (let i = 0; i < arr.length; i += size) out.push(arr.slice(i, i + size));
  return out;
}

// 重複していた applySentFlagIfNeeded_（boolean版）は削除しました。日時記録版のみを使用します。

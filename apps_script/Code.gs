/**
 * はやれき 入力アプリ（Google Apps Script）
 *
 * 1. Google スプレッドシートの「拡張機能 → Apps Script」を開く
 * 2. このファイルを Code.gs に、index.html を HTML ファイル「index」に貼り付ける
 * 3. 関数 setup を一度だけ実行する（シートと見出しが作られる）
 * 4. 「デプロイ → 新しいデプロイ → ウェブアプリ」で公開し、URL をスマホで開く
 */

const SHEET = { rec: '記録', plots: '区画', ferts: '肥料', pests: '農薬', works: '作業' };

const HEAD = {
  rec:   ['id', 'event_id', '登録日時', '日付', '区画', '分類', '項目', '使用量', '単位', '水量L', '開花段', '収穫段', 'メモ'],
  plots: ['区画名', '面積a', 'ハウス数'],
  ferts: ['資材名', '全N%', '化学N%', 'リン酸%', 'カリ%'],
  pests: ['薬剤名', '分類', '使用回数上限', 'カウント', '計算方法', '最小倍率', '最大倍率', '単位'],
  works: ['作業名'],
};

// 初回のみの見本。setup 後に自分の内容へ書き換える
const SAMPLE = {
  plots: [['第1', 6, 3], ['第2', 6, 3], ['第3', 4, 2]],
  ferts: [['液肥A', 10, 9.7, 4, 6], ['カリ資材B', 5, 5, 0, 49]],
  pests: [['殺菌剤A', '殺菌剤', 3, 1, '希釈', 2000, '', 'ml'],
          ['展着剤C', '展着剤', 0, 0, '希釈', 2000, '', 'ml'],
          ['粒剤D', '殺虫剤', 1, 1, '直', '', '', 'g']],
  works: [['芽かき'], ['誘引'], ['葉かき'], ['摘果'], ['収穫']],
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('はやれき')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

function setup() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(SHEET).forEach(k => {
    let sh = ss.getSheetByName(SHEET[k]);
    if (!sh) sh = ss.insertSheet(SHEET[k]);
    if (sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, HEAD[k].length).setValues([HEAD[k]]).setFontWeight('bold');
      sh.setFrozenRows(1);
      if (SAMPLE[k]) sh.getRange(2, 1, SAMPLE[k].length, HEAD[k].length).setValues(SAMPLE[k]);
    }
  });
  const rec = ss.getSheetByName(SHEET.rec);
  rec.getRange('D:D').setNumberFormat('yyyy/mm/dd');
  rec.getRange('C:C').setNumberFormat('yyyy/mm/dd hh:mm');
}

function table_(name) {
  const sh = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sh || sh.getLastRow() < 2) return [];
  const v = sh.getDataRange().getValues();
  const h = v.shift();
  return v.filter(r => r[0] !== '').map(r => Object.fromEntries(h.map((k, i) => [k, r[i]])));
}

function getMasters() {
  return {
    plots: table_(SHEET.plots), ferts: table_(SHEET.ferts),
    pests: table_(SHEET.pests), works: table_(SHEET.works),
    recent: recent_(8),
  };
}

/**
 * p = { date:'2026-09-16', category:'防除', plots:['第1','第2'],
 *       items:[{name:'殺菌剤A', qty:250, unit:'ml'}], water:500,
 *       flower:'', harvest:'', note:'' }
 * 区画×項目ごとに1行ずつ保存する。同じ保存操作の行は event_id でつながる。
 */
function saveRecord(p) {
  if (!p || !p.date || !p.category || !p.plots || !p.plots.length || !p.items || !p.items.length) {
    throw new Error('日付・分類・区画・項目を入れてください');
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET.rec);
    const ev = Utilities.getUuid().slice(0, 8);
    const now = new Date();
    const day = new Date(p.date + 'T00:00:00');
    const n = v => (v === '' || v === null || v === undefined) ? '' : Number(v);
    const rows = [];
    p.items.forEach(it => p.plots.forEach(plot => rows.push([
      Utilities.getUuid().slice(0, 8), ev, now, day, plot, p.category, it.name,
      n(it.qty), it.unit || '', n(p.water), n(p.flower), n(p.harvest), p.note || '',
    ])));
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, HEAD.rec.length).setValues(rows);
    return { saved: rows.length, recent: recent_(8) };
  } finally {
    lock.releaseLock();
  }
}

/** 入力の取り消し（同じ保存操作の行をまとめて消す） */
function deleteEvent(ev) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET.rec);
    const v = sh.getRange(1, 2, sh.getLastRow(), 1).getValues();
    for (let i = v.length - 1; i >= 1; i--) {
      if (String(v[i][0]) === String(ev)) sh.deleteRow(i + 1);
    }
    return recent_(8);
  } finally {
    lock.releaseLock();
  }
}

function recent_(n) {
  const groups = {};
  table_(SHEET.rec).forEach(r => {
    const g = groups[r.event_id] || (groups[r.event_id] = {
      ev: String(r.event_id), at: new Date(r['登録日時']).getTime(), date: r['日付'],
      cat: r['分類'], items: [], plots: [] });
    if (!g.items.includes(r['項目'])) g.items.push(r['項目']);
    if (!g.plots.includes(r['区画'])) g.plots.push(r['区画']);
  });
  return Object.values(groups).sort((a, b) => b.at - a.at).slice(0, n).map(g => ({
    ev: g.ev, cat: g.cat, items: g.items.join('・'), plots: g.plots.join('・'),
    date: Utilities.formatDate(new Date(g.date), 'Asia/Tokyo', 'M/d'),
  }));
}

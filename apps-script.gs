/**
 * 竹林小隊練功坊：雲端紀錄（Google 試算表）
 * 工作表「小勇士」「對決」「對決結果」「動態」會在第一次寫入時自動建立。
 *
 * 安裝：建立新的 Google 試算表 → 擴充功能 → Apps Script → 整段貼上取代、存檔
 *       → 部署 → 新增部署作業 → 類型選「網頁應用程式」
 *         執行身分：我；誰可以存取：所有人 → 部署，複製網址（…/exec）。
 * 更新：貼上新程式碼存檔 → 部署 → 管理部署作業 → 編輯（鉛筆）→ 版本：新版本 → 部署，網址不變。
 */
const FAMILY_CODE = 'PUT-YOUR-FAMILY-CODE-HERE';   // 家庭通關密語：網頁也要輸入同一組才能讀寫

const SHEETS = {
  player:    { name: '小勇士',   headers: ['ID', '名字', '等級', '經驗值', '星星', '連續天數', '最後練習', '更新時間', '資料'], dataCol: 9 },
  challenge: { name: '對決',     headers: ['ID', '發起人', '對象', '項目', '建立時間', '更新時間', '資料'], dataCol: 7 },
  result:    { name: '對決結果', headers: ['ID', '對決ID', '玩家', '成績', '更新時間', '資料'], dataCol: 6 },
  feed:      { name: '動態',     headers: ['ID', '玩家', '內容', '時間', '資料'], dataCol: 5 }
};

function out_(s) { return ContentService.createTextOutput(s); }
function sheet_(k) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(), c = SHEETS[k];
  let sh = ss.getSheetByName(c.name);
  if (!sh) { sh = ss.insertSheet(c.name); sh.appendRow(c.headers); sh.setFrozenRows(1); }
  return sh;
}
function clean_(v) { v = String(v == null ? '' : v).slice(0, 500); return /^[=+\-@]/.test(v) ? "'" + v : v; }
function okId_(id) { return /^[A-Za-z0-9_]{3,80}$/.test(String(id || '')); }
function findRow_(sh, id) {
  const n = sh.getLastRow(); if (n < 2) return 0;
  const ids = sh.getRange(2, 1, n - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return 0;
}
function lv_(xp) { xp = Number(xp) || 0; let l = 1, need = 100; while (xp >= need) { xp -= need; l++; need = 100 + (l - 1) * 30; } return l; }
function stars_(p) { let s = 0; const st = p.stars || {}; for (const k in st) { if (/^boss-/.test(k)) continue; (st[k] || []).forEach(v => s += Number(v) || 0); } return s; }
function nameOf_(id) { const sh = sheet_('player'), r = findRow_(sh, id); return r ? String(sh.getRange(r, 2).getValue()) : id; }

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const d = JSON.parse(e.postData.contents);
    if (String(d.code) !== FAMILY_CODE) return out_('bad code');
    const kind = d.kind, id = String(d.id || ''), data = d.data || {};
    if (!okId_(id)) return out_('bad id');
    if (kind === 'delplayer') { const sh = sheet_('player'), r = findRow_(sh, id); if (r) sh.deleteRow(r); return out_('ok'); }
    if (!SHEETS[kind]) return out_('bad kind');
    const json = JSON.stringify(data);
    if (json.length > 45000) return out_('too big');
    const sh = sheet_(kind), r = findRow_(sh, id), t = new Date();
    let row;
    if (kind === 'player') {
      if (r) { let old = {}; try { old = JSON.parse(sh.getRange(r, 9).getValue() || '{}'); } catch (x) {} if ((Number(old.rev) || 0) > (Number(data.rev) || 0)) return out_('older'); }
      row = [id, clean_(data.name), lv_(data.xp), Number(data.xp) || 0, stars_(data), Number(data.streak) || 0, clean_(data.lastDay), t, json];
    } else {
      if (r) return out_('exists');
      if (kind === 'challenge') row = [id, clean_(nameOf_(data.from)), clean_((data.to || []).map(nameOf_).join('、')), clean_(data.kind === 'arcade' ? '九九快打' : (data.node + ' Lv' + data.lv)), new Date(Number(data.at) || t), t, json];
      else if (kind === 'result') row = [id, clean_(data.cid), clean_(nameOf_(data.pid)), data.count != null ? Number(data.count) + ' 題' : Number(data.pts) + ' 分', t, json];
      else row = [id, clean_(nameOf_(data.pid)), clean_((data.icon || '') + ' ' + (data.msg || '')), t, json];
    }
    if (r) sh.getRange(r, 1, 1, row.length).setValues([row]); else sh.appendRow(row);
    if (kind === 'feed') { const n = sh.getLastRow(); if (n > 301) sh.deleteRows(2, n - 301); }
    return out_('ok');
  } finally {
    lock.releaseLock();
  }
}

/* 讀取：?action=all&code=密語&cb=函式名 → 全部角色、近 10 天的對決與結果、最新 40 則動態（JSONP） */
function doGet(e) {
  const p = (e && e.parameter) || {}, cb = String(p.cb || '');
  if (!/^[A-Za-z_$][\w$]{0,40}$/.test(cb)) return out_('竹林小隊練功坊紀錄服務運作中');
  let res;
  if (String(p.code) !== FAMILY_CODE) res = { error: 'code' };
  else {
    const read = (k, filter) => {
      const sh = sheet_(k), n = sh.getLastRow(); if (n < 2) return [];
      return sh.getRange(2, SHEETS[k].dataCol, n - 1, 1).getValues()
        .map(r => { try { return JSON.parse(r[0]); } catch (x) { return null; } })
        .filter(x => x && (!filter || filter(x)));
    };
    const cut = Date.now() - 10 * 86400000;
    const challenges = read('challenge', c => Number(c.at) > cut), cids = {};
    challenges.forEach(c => cids[c.id] = 1);
    res = { players: read('player'), challenges: challenges, results: read('result', r => cids[r.cid]), feed: read('feed').slice(-40) };
  }
  return ContentService.createTextOutput(cb + '(' + JSON.stringify(res) + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
}

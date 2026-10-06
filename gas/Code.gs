const SS = SpreadsheetApp.getActiveSpreadsheet();

const COLS = ['id','requester','name','qty','cat','place','url','memo','img','created','boughtAt','heldAt','requestedQty','boughtQty','garden','folder','partial','imgFull'];
const USUAL_COLS = ['id','name','qty','cat','place','memo','img','garden','folder','order','imgFull'];

function getSheet(name) {
  let sh = SS.getSheetByName(name);
  if (!sh) {
    sh = SS.insertSheet(name);
    if (['pending','done','hold'].includes(name)) {
      sh.appendRow(COLS);
      sh.setFrozenRows(1);
    }
    if (name === 'masters') {
      sh.appendRow(['type','value','garden','order']);
      sh.setFrozenRows(1);
    }
    if (name === 'folders') {
      sh.appendRow(['id','name','garden','created','order']);
      sh.setFrozenRows(1);
    }
    if (name === 'usualItems') {
      sh.appendRow(USUAL_COLS);
      sh.setFrozenRows(1);
    }
    if (name === 'images') {
      sh.appendRow(['imgId','seq','chunk']);
      sh.setFrozenRows(1);
    }
  }
  return sh;
}

function doGet(e) {
  const action = e.parameter.action;
  const garden = e.parameter.garden || '';
  let result;
  if (action === 'getAll') {
    result = {
      pending: sheetToArray('pending', garden),
      done:    sheetToArray('done',    garden),
      hold:    sheetToArray('hold',    garden),
      masters: getMasters(garden),
      folders: getFolders(garden),
      usualItems: getUsualItems(garden)
    };
  } else if (action === 'getImage') {
    result = getImage(e.parameter.id);
  } else if (action === 'getItemImg') {
    result = getItemImg(e.parameter.sheet, e.parameter.id);
  } else {
    result = { error: 'unknown action' };
  }
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    let result = { ok: true };

    if (action === 'addPending') {
      appendItem('pending', data.item);
      saveMaster('requester', data.item.requester, data.item.garden||'');
      if (data.item.place) saveMaster('place', data.item.place, data.item.garden||'');

    } else if (action === 'moveToDone') {
      const extra = { boughtAt: data.boughtAt };
      if (data.requestedQty) extra.requestedQty = data.requestedQty;
      if (data.boughtQty)    extra.boughtQty    = data.boughtQty;
      moveRow('pending', 'done', data.id, extra);

    } else if (action === 'partialBuy') {
      const src = findRow('pending', data.id);
      if (src) {
        src.boughtAt     = data.boughtAt;
        src.requestedQty = data.requestedQty;
        src.boughtQty    = data.boughtQty;
        src.qty          = data.boughtQty;
        appendItem('done', src);
        updateItemFields('pending', data.id, { qty: data.remainingQty, partial: '1' });
      }

    } else if (action === 'moveToHold') {
      moveRow('pending', 'hold', data.id, { heldAt: data.heldAt });

    } else if (action === 'holdToPublic') {
      moveRow('hold', 'pending', data.id, {});

    } else if (action === 'doneToPublic') {
      moveRow('done', 'pending', data.id, {});

    } else if (action === 'deletePending') {
      deleteItemWithImage('pending', data.id);

    } else if (action === 'deleteDone') {
      deleteItemWithImage('done', data.id);

    } else if (action === 'deleteHold') {
      deleteItemWithImage('hold', data.id);

    } else if (action === 'rebuy') {
      const src = findRow('done', data.id);
      if (src) {
        src.id  = String(Date.now());
        src.qty = '';
        src.requestedQty = '';
        src.boughtQty    = '';
        src.created = new Date().toLocaleDateString('ja-JP');
        delete src.boughtAt;
        appendItem('pending', src);
      }

    } else if (action === 'updateItem') {
      const sheet = data.sheet || 'pending';
      updateItemWithImage(sheet, data.id, data.fields);
      if (data.fields && data.fields.place) {
        const item = findRow(sheet, data.id);
        saveMaster('place', data.fields.place, item?.garden||'');
      }

    } else if (action === 'updateCat') {
      updateItemFields('pending', data.id, { cat: data.cat });

    } else if (action === 'deleteMaster') {
      deleteMasterRow(data.masterType, data.value, data.garden||'');

    } else if (action === 'saveMasterDirect') {
      saveMaster(data.masterType, data.value, data.garden||'');

    } else if (action === 'reorderMaster') {
      reorderMaster(data.masterType, data.garden||'', data.values);

    } else if (action === 'addFolder') {
      addFolder(data.name, data.garden||'');

    } else if (action === 'deleteFolder') {
      deleteFolder(data.folderId, data.garden||'');
      resetFolderItems('pending', data.folderId);
      resetFolderItems('done', data.folderId);
      resetFolderItems('hold', data.folderId);
      resetFolderItems('usualItems', data.folderId);

    } else if (action === 'reorderFolders') {
      reorderFolders(data.garden||'', data.folderIds);

    } else if (action === 'moveToFolder') {
      updateItemFields(data.sheet||'pending', data.id, { folder: data.folderId });

    } else if (action === 'addUsualItem') {
      appendUsualItem(data.item);

    } else if (action === 'deleteUsualItem') {
      deleteItemWithImage('usualItems', data.id);

    } else if (action === 'updateUsualItem') {
      updateItemWithImage('usualItems', data.id, data.fields);

    } else if (action === 'saveImage') {
      result = saveImage(data.id, data.data);

    } else if (action === 'deleteImage') {
      result = deleteImage(data.id);

    } else {
      result = { error: 'unknown action' };
    }

    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function sheetToArray(sheetName, garden) {
  const sh = getSheet(sheetName);
  const rows = sh.getDataRange().getValues();
  if (rows.length <= 1) return [];
  const headers = rows[0];
  const gardenCol = headers.indexOf('garden');
  return rows.slice(1)
    .filter(r => {
      if (!garden) return true;
      const g = gardenCol >= 0 ? String(r[gardenCol] ?? '') : '';
      return g === garden ;
    })
    .map(r => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = String(r[i] ?? ''); });
      return slimImg(obj);
    });
}

// 旧方式の大きな画像は一覧に含めない（起動を軽くするため）。拡大時に getItemImg で個別に取得する
const THUMB_MAX = 12000;
function slimImg(obj) {
  if (obj.img && obj.img.length > THUMB_MAX) {
    obj.img = '';
    obj.imgBig = '1';
  }
  return obj;
}

function getItemImg(sheetName, id) {
  if (!IMG_REF_SHEETS.includes(sheetName)) return { ok: false, error: 'bad sheet' };
  const item = findRow(sheetName, id);
  if (!item || !item.img) return { ok: false, error: 'not found' };
  return { ok: true, data: item.img };
}

// 1行目の見出しに足りない列があれば右端に追加し、見出し配列を返す
function ensureHeaders(sh, cols) {
  const lastCol = sh.getLastColumn();
  const headers = lastCol > 0 ? sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String) : [];
  const missing = cols.filter(c => !headers.includes(c));
  if (missing.length) {
    sh.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]);
    headers.push(...missing);
  }
  return headers;
}

// 見出し名に合わせて1行追加（列の並び順に依存しない）
function appendByHeaders(sh, cols, item) {
  const headers = ensureHeaders(sh, cols);
  sh.appendRow(headers.map(h => item[h] ?? ''));
}

function appendItem(sheetName, item) {
  const sh = getSheet(sheetName);
  if (!item.id) item.id = String(Date.now());
  appendByHeaders(sh, COLS, item);
}

function findRow(sheetName, id) {
  const sh = getSheet(sheetName);
  const rows = sh.getDataRange().getValues();
  const headers = rows[0];
  const idCol = headers.indexOf('id');
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][idCol]) === String(id)) {
      const obj = {};
      headers.forEach((h, j) => { obj[h] = String(rows[i][j] ?? ''); });
      return obj;
    }
  }
  return null;
}

function deleteRow(sheetName, id) {
  const sh = getSheet(sheetName);
  const rows = sh.getDataRange().getValues();
  const idCol = rows[0].indexOf('id');
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][idCol]) === String(id)) {
      sh.deleteRow(i + 1);
      break;
    }
  }
}

// アイテムを削除し、どこからも使われなくなった画像チャンクも削除
function deleteItemWithImage(sheetName, id) {
  const item = findRow(sheetName, id);
  deleteRow(sheetName, id);
  if (item && item.imgFull) deleteImage(item.imgFull);
}

// 更新で画像が差し替え・削除されたら、古い画像チャンクを削除
function updateItemWithImage(sheetName, id, fields) {
  const before = (fields && 'imgFull' in fields) ? findRow(sheetName, id) : null;
  updateItemFields(sheetName, id, fields);
  if (before && before.imgFull && before.imgFull !== String(fields.imgFull || '')) {
    deleteImage(before.imgFull);
  }
}

function moveRow(fromSheet, toSheet, id, extra = {}) {
  const item = findRow(fromSheet, id);
  if (!item) return;
  Object.assign(item, extra);
  appendItem(toSheet, item);
  deleteRow(fromSheet, id);
}

function updateItemFields(sheetName, id, fields) {
  const sh = getSheet(sheetName);
  if (fields && 'imgFull' in fields) ensureHeaders(sh, ['imgFull']);
  const rows = sh.getDataRange().getValues();
  const headers = rows[0];
  const idCol = headers.indexOf('id');
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][idCol]) === String(id)) {
      Object.entries(fields).forEach(([key, val]) => {
        const col = headers.indexOf(key);
        if (col >= 0) sh.getRange(i + 1, col + 1).setValue(val);
      });
      break;
    }
  }
}

function getMasters(garden) {
  const sh = getSheet('masters');
  const rows = sh.getDataRange().getValues();
  if (rows.length <= 1) return { requesters: [], places: [] };
  const headers = rows[0];
  const gardenCol = headers.indexOf('garden');
  const orderCol = headers.indexOf('order');
  const requesterEntries = [], placeEntries = [];
  rows.slice(1).forEach(r => {
    const g = gardenCol >= 0 ? String(r[gardenCol] ?? '') : '';
    if (garden && g !== garden && g !== '') return;
    const ord = orderCol >= 0 && r[orderCol] !== '' ? Number(r[orderCol]) : 999999;
    if (r[0] === 'requester' && r[1]) requesterEntries.push({ value: String(r[1]), order: ord });
    if (r[0] === 'place'     && r[1]) placeEntries.push({ value: String(r[1]), order: ord });
  });
  const dedupe = (entries) => {
    const seen = new Map();
    entries.forEach(e => {
      if (!seen.has(e.value) || seen.get(e.value) > e.order) seen.set(e.value, e.order);
    });
    return [...seen.entries()]
      .sort((a, b) => a[1] - b[1])
      .map(([value]) => value);
  };
  return { requesters: dedupe(requesterEntries), places: dedupe(placeEntries) };
}

function saveMaster(type, value, garden) {
  if (!value) return;
  const sh = getSheet('masters');
  const rows = sh.getDataRange().getValues();
  const headers = rows[0];
  const gardenCol = headers.indexOf('garden');
  const orderCol = headers.indexOf('order');
  let maxOrder = -1;
  let exists = false;
  rows.slice(1).forEach(r => {
    const g = gardenCol >= 0 ? String(r[gardenCol] ?? '') : '';
    if (g !== (garden||'')) return;
    if (r[0] === type) {
      const ord = orderCol >= 0 && r[orderCol] !== '' ? Number(r[orderCol]) : 0;
      if (ord > maxOrder) maxOrder = ord;
      if (String(r[1]) === String(value)) exists = true;
    }
  });
  if (!exists) sh.appendRow([type, value, garden||'', maxOrder + 1]);
}

function deleteMasterRow(type, value, garden) {
  const sh = getSheet('masters');
  const rows = sh.getDataRange().getValues();
  const headers = rows[0];
  const gardenCol = headers.indexOf('garden');
  for (let i = rows.length - 1; i >= 1; i--) {
    const g = gardenCol >= 0 ? String(rows[i][gardenCol] ?? '') : '';
    if (String(rows[i][0]) === String(type) && String(rows[i][1]) === String(value) && g === (garden||'')) {
      sh.deleteRow(i + 1);
    }
  }
}

function reorderMaster(type, garden, values) {
  const sh = getSheet('masters');
  const rows = sh.getDataRange().getValues();
  const headers = rows[0];
  const gardenCol = headers.indexOf('garden');
  let orderCol = headers.indexOf('order');
  if (orderCol < 0) {
    sh.getRange(1, headers.length + 1).setValue('order');
    orderCol = headers.length;
  }
  const orderMap = {};
  values.forEach((v, idx) => { orderMap[v] = idx; });
  for (let i = 1; i < rows.length; i++) {
    const g = gardenCol >= 0 ? String(rows[i][gardenCol] ?? '') : '';
    if (g !== (garden||'')) continue;
    if (String(rows[i][0]) !== type) continue;
    const val = String(rows[i][1]);
    if (val in orderMap) {
      sh.getRange(i + 1, orderCol + 1).setValue(orderMap[val]);
    }
  }
}

function getFolders(garden) {
  const sh = getSheet('folders');
  const rows = sh.getDataRange().getValues();
  if (rows.length <= 1) return [];
  const headers = rows[0];
  const gardenCol = headers.indexOf('garden');
  const orderCol = headers.indexOf('order');
  const list = rows.slice(1)
    .filter(r => {
      const g = gardenCol >= 0 ? String(r[gardenCol] ?? '') : '';
      return !garden || g === garden || g === '';
    })
    .map(r => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = String(r[i] ?? ''); });
      obj._order = orderCol >= 0 && r[orderCol] !== '' ? Number(r[orderCol]) : 999999;
      return obj;
    });
  list.sort((a, b) => a._order - b._order);
  list.forEach(o => delete o._order);
  return list;
}

function addFolder(name, garden) {
  const sh = getSheet('folders');
  const rows = sh.getDataRange().getValues();
  const headers = rows[0];
  const gardenCol = headers.indexOf('garden');
  const orderCol = headers.indexOf('order');
  let maxOrder = -1;
  rows.slice(1).forEach(r => {
    const g = gardenCol >= 0 ? String(r[gardenCol] ?? '') : '';
    if (g !== (garden||'')) return;
    const ord = orderCol >= 0 && r[orderCol] !== '' ? Number(r[orderCol]) : 0;
    if (ord > maxOrder) maxOrder = ord;
  });
  const id = String(Date.now());
  sh.appendRow([id, name, garden, new Date().toLocaleDateString('ja-JP'), maxOrder + 1]);
}

function deleteFolder(folderId, garden) {
  const sh = getSheet('folders');
  const rows = sh.getDataRange().getValues();
  const idCol = rows[0].indexOf('id');
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][idCol]) === String(folderId)) {
      sh.deleteRow(i + 1);
      break;
    }
  }
}

function reorderFolders(garden, folderIds) {
  const sh = getSheet('folders');
  const rows = sh.getDataRange().getValues();
  const headers = rows[0];
  const gardenCol = headers.indexOf('garden');
  const idCol = headers.indexOf('id');
  let orderCol = headers.indexOf('order');
  if (orderCol < 0) {
    sh.getRange(1, headers.length + 1).setValue('order');
    orderCol = headers.length;
  }
  const orderMap = {};
  folderIds.forEach((id, idx) => { orderMap[id] = idx; });
  for (let i = 1; i < rows.length; i++) {
    const g = gardenCol >= 0 ? String(rows[i][gardenCol] ?? '') : '';
    if (g !== (garden||'')) continue;
    const id = String(rows[i][idCol]);
    if (id in orderMap) {
      sh.getRange(i + 1, orderCol + 1).setValue(orderMap[id]);
    }
  }
}

function resetFolderItems(sheetName, folderId) {
  const sh = getSheet(sheetName);
  const rows = sh.getDataRange().getValues();
  const headers = rows[0];
  const folderCol = headers.indexOf('folder');
  if (folderCol < 0) return;
  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][folderCol]) === String(folderId)) {
      sh.getRange(i + 1, folderCol + 1).setValue('');
    }
  }
}

/* ── いつものリクエスト（定番アイテム） ── */
function getUsualItems(garden) {
  const sh = getSheet('usualItems');
  const rows = sh.getDataRange().getValues();
  if (rows.length <= 1) return [];
  const headers = rows[0];
  const gardenCol = headers.indexOf('garden');
  const orderCol = headers.indexOf('order');
  const list = rows.slice(1)
    .filter(r => {
      const g = gardenCol >= 0 ? String(r[gardenCol] ?? '') : '';
      return !garden || g === garden || g === '';
    })
    .map(r => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = String(r[i] ?? ''); });
      obj._order = orderCol >= 0 && r[orderCol] !== '' ? Number(r[orderCol]) : 999999;
      return slimImg(obj);
    });
  list.sort((a, b) => a._order - b._order);
  list.forEach(o => delete o._order);
  return list;
}

function appendUsualItem(item) {
  const sh = getSheet('usualItems');
  if (!item.id) item.id = String(Date.now());
  appendByHeaders(sh, USUAL_COLS, item);
}

/* ── 画像の分割保存（imagesシート：imgId | seq | chunk） ── */
const IMG_CHUNK = 40000;
const IMG_REF_SHEETS = ['pending', 'done', 'hold', 'usualItems'];

function imgRows(sh, id) {
  if (!id || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 1)
    .createTextFinder(String(id)).matchEntireCell(true).findAll()
    .map(r => r.getRow());
}

function getImage(id) {
  const sh = getSheet('images');
  const rows = imgRows(sh, id);
  if (!rows.length) return { ok: false, error: 'not found' };
  const parts = rows.map(r => sh.getRange(r, 2, 1, 2).getValues()[0])
    .sort((a, b) => Number(a[0]) - Number(b[0]))
    .map(v => String(v[1]).replace(/^c:/, '')); // 保護用プレフィックスを外す
  return { ok: true, data: parts.join('') };
}

function saveImage(id, data) {
  if (!id || !data) return { ok: false, error: 'id/data required' };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getSheet('images');
    deleteImageRows(sh, id); // 同じIDが残っていたら上書き
    const rows = [];
    for (let i = 0, seq = 0; i < data.length; i += IMG_CHUNK, seq++) {
      // base64が「+」「=」で始まると数式扱いされるため「c:」を付けて保存
      rows.push([String(id), seq, 'c:' + data.slice(i, i + IMG_CHUNK)]);
    }
    sh.getRange(sh.getLastRow() + 1, 1, rows.length, 3).setValues(rows);
    return { ok: true, id: id, chunks: rows.length };
  } finally {
    lock.releaseLock();
  }
}

// どのアイテムからも参照されていなければ画像チャンクを削除
function deleteImage(id) {
  if (!id) return { ok: false, error: 'id required' };
  if (isImageReferenced(id)) return { ok: true, deleted: false, reason: 'still referenced' };
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const n = deleteImageRows(getSheet('images'), id);
    return { ok: true, deleted: n > 0 };
  } finally {
    lock.releaseLock();
  }
}

function deleteImageRows(sh, id) {
  const rows = imgRows(sh, id).sort((a, b) => b - a);
  let i = 0;
  while (i < rows.length) { // 連続した行はまとめて削除
    let start = rows[i], count = 1;
    while (i + count < rows.length && rows[i + count] === start - 1) { start--; count++; }
    sh.deleteRows(start, count);
    i += count;
  }
  return rows.length;
}

function isImageReferenced(id) {
  return IMG_REF_SHEETS.some(name => {
    const sh = SS.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) return false;
    const col = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].indexOf('imgFull') + 1;
    if (!col) return false;
    return !!sh.getRange(2, col, sh.getLastRow() - 1, 1)
      .createTextFinder(String(id)).matchEntireCell(true).findNext();
  });
}

// 手動実行用（1回だけ）：旧方式の大きな img を images シートへ移し、セルを空にする
// 移した画像は imgFull から拡大表示でき、サムネイルは初めて拡大したときにアプリ側で作り直される
function migrateLegacyImages() {
  const imgSh = getSheet('images');
  const existing = new Set(imgSh.getLastRow() < 2 ? [] :
    imgSh.getRange(2, 1, imgSh.getLastRow() - 1, 1).getValues().map(r => String(r[0])));
  const chunks = [];
  const plans = [];
  IMG_REF_SHEETS.forEach(name => {
    const sh = SS.getSheetByName(name);
    if (!sh || sh.getLastRow() < 2) return;
    const headers = ensureHeaders(sh, ['imgFull']);
    const idCol = headers.indexOf('id'), imgCol = headers.indexOf('img'), fullCol = headers.indexOf('imgFull');
    if (idCol < 0 || imgCol < 0) return;
    const n = sh.getLastRow() - 1;
    const ids = sh.getRange(2, idCol + 1, n, 1).getValues().map(r => String(r[0]));
    const imgs = sh.getRange(2, imgCol + 1, n, 1).getValues();
    const fulls = sh.getRange(2, fullCol + 1, n, 1).getValues();
    let count = 0;
    for (let i = 0; i < n; i++) {
      const img = String(imgs[i][0] || '');
      if (img.length <= THUMB_MAX) continue;
      if (!String(fulls[i][0] || '')) {
        // 同じ画像（再購入でコピーされた分など）は1つのIDにまとめる
        const id = 'lg' + Utilities.base64EncodeWebSafe(
          Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, img)).replace(/=+$/, '');
        if (!existing.has(id)) {
          existing.add(id);
          for (let p = 0, seq = 0; p < img.length; p += IMG_CHUNK, seq++) {
            chunks.push([id, seq, 'c:' + img.slice(p, p + IMG_CHUNK)]);
          }
        }
        fulls[i][0] = id;
      }
      imgs[i][0] = '';
      count++;
    }
    if (count) plans.push({ name, sh, ids, imgs, fulls, imgCol, fullCol, count });
  });

  // 先に画像チャンクを書き込んでから、元のセルを差し替える
  for (let i = 0; i < chunks.length; i += 50) {
    const part = chunks.slice(i, i + 50);
    imgSh.getRange(imgSh.getLastRow() + 1, 1, part.length, 3).setValues(part);
  }
  plans.forEach(p => {
    const n = p.ids.length;
    // 実行中に行の追加・削除があった場合はずれるので、その表は書き込まない
    const idCol = ensureHeaders(p.sh, ['imgFull']).indexOf('id');
    const nowIds = p.sh.getRange(2, idCol + 1, n, 1).getValues().map(r => String(r[0]));
    if (nowIds.join('\n') !== p.ids.join('\n')) {
      Logger.log(p.name + '：実行中に行が変わったためスキップしました。もう一度実行してください');
      return;
    }
    p.sh.getRange(2, p.fullCol + 1, n, 1).setValues(p.fulls);
    p.sh.getRange(2, p.imgCol + 1, n, 1).setValues(p.imgs);
    Logger.log(p.name + '：' + p.count + '件の画像を移しました');
  });
  Logger.log('画像チャンク ' + chunks.length + '行を追加しました');
}

// 手動実行用：どこからも参照されていない画像チャンクを一括削除
function cleanupOrphanImages() {
  const sh = getSheet('images');
  if (sh.getLastRow() < 2) return;
  const ids = [...new Set(sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0])))];
  ids.filter(id => id && !isImageReferenced(id)).forEach(id => deleteImageRows(sh, id));
}

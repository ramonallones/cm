const SHEETS = {
  MASTER_PRODUK: ['SKU','Nama Produk','Kategori','Harga Modal','Harga Jual','Stok','Aktif','Updated At','Brand'],
  ORDER: ['Order ID','Tanggal','Nama Customer','No WhatsApp','Alamat','Ekspedisi','Ongkir','Diskon','Subtotal','Total','Status Pembayaran','Status Order','Catatan'],
  ORDER_ITEM: ['Order ID','SKU','Nama Produk','Harga','Qty','Total'],
  MUTASI_STOK: ['Tanggal','SKU','Jenis','Qty','Referensi','Keterangan'],
  PELANGGAN: ['No WhatsApp','Nama Customer','Alamat','Order Terakhir'],
  REKAP_HARIAN: ['Tanggal','Jumlah Order','Omzet'],
  REKAP_BULANAN: ['Bulan','Jumlah Order','Omzet'],
  SETTING: ['Key','Value']
};
const DEFAULT_APP_PIN = '123456';
const AUTH_TTL_SECONDS = 21600;

function doGet(e) {
  try {
    setupSheets_();
    const action = String(e.parameter.action || '');
    requireAuth_(e.parameter.token);
    if (action === 'checkAuth') return json_({ success: true });
    if (action === 'getDashboard') return json_(getDashboard_());
    if (action === 'getProducts') return json_({ success: true, products: rows_('MASTER_PRODUK') });
    if (action === 'getOrders') return json_({ success: true, orders: getOrders_() });
    if (action === 'getMutasiStock') return json_({ success: true, data: rows_('MUTASI_STOK') });
    if (action === 'getPelanggan') return json_({ success: true, data: rows_('PELANGGAN') });
    if (action === 'setup') return json_(setupSheets_());
    throw new Error('Action GET tidak dikenal: ' + action);
  } catch (error) {
    return json_({ success: false, message: error.message });
  }
}

function doPost(e) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    setupSheets_();
    const action = String(e.parameter.action || '');
    const payload = JSON.parse(e.postData && e.postData.contents || '{}');
    if (action === 'authenticate') return json_(authenticate_(payload.pin));
    requireAuth_(payload.token);
    const handlers = {
      addProduct: () => addProduct_(payload),
      updateProduct: () => updateProduct_(payload),
      nonaktifProduct: () => nonaktifProduct_(payload),
      addStock: () => addStock_(payload),
      addBonus: () => addBonus_(payload),
      addBonuses: () => addBonuses_(payload),
      saveOrder: () => saveOrder_(payload),
      deleteOrder: () => deleteOrders_([payload.orderId]),
      deleteOrders: () => deleteOrders_(payload.orderIds || [])
    };
    if (!handlers[action]) throw new Error('Action POST tidak dikenal: ' + action);
    return json_(handlers[action]());
  } catch (error) {
    return json_({ success: false, message: error.message });
  } finally {
    lock.releaseLock();
  }
}

function setupSheets_() {
  const ss = SpreadsheetApp.getActive();
  Object.keys(SHEETS).forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    if (sheet.getLastRow() === 0) sheet.appendRow(SHEETS[name]);
    else {
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      SHEETS[name].forEach(header => {
        if (!headers.includes(header)) {
          sheet.getRange(1, sheet.getLastColumn() + 1).setValue(header);
          headers.push(header);
        }
      });
    }
    sheet.setFrozenRows(1);
  });
  ensureDefaultPin_();
  return { success: true, sheets: Object.keys(SHEETS) };
}

function ensureDefaultPin_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('SETTING');
  const values = sheet.getDataRange().getValues();
  if (!values.slice(1).some(row => String(row[0]) === 'APP_PIN')) sheet.appendRow(['APP_PIN', DEFAULT_APP_PIN]);
}

function settingValue_(key) {
  const values = SpreadsheetApp.getActive().getSheetByName('SETTING').getDataRange().getValues();
  for (let row = 1; row < values.length; row++) {
    if (String(values[row][0]) === String(key)) return String(values[row][1]);
  }
  return '';
}

function authenticate_(pin) {
  if (String(pin || '') !== settingValue_('APP_PIN')) throw new Error('PIN salah.');
  const token = Utilities.getUuid() + Utilities.getUuid();
  CacheService.getScriptCache().put('auth_' + token, '1', AUTH_TTL_SECONDS);
  return { success: true, token: token, expiresIn: AUTH_TTL_SECONDS };
}

function requireAuth_(token) {
  if (!token || CacheService.getScriptCache().get('auth_' + token) !== '1') throw new Error('Sesi tidak valid atau sudah berakhir. Silakan masukkan PIN lagi.');
}

function rows_(name) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const values = sheet.getDataRange().getValues();
  const headers = values.shift();
  return values.filter(row => row.some(value => value !== '')).map(row => headers.reduce((out, key, i) => {
    out[key] = row[i] instanceof Date ? row[i].toISOString() : row[i];
    return out;
  }, {}));
}

function getDashboard_() {
  const products = rows_('MASTER_PRODUK');
  const orders = rows_('ORDER');
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const todayOrders = orders.filter(order => String(order.Tanggal).slice(0, 10) === today);
  return {
    success: true,
    dashboard: {
      totalProduk: products.length,
      produkAktif: products.filter(product => Number(product.Aktif) !== 0).length,
      totalStok: products.reduce((sum, product) => sum + Number(product.Stok || 0), 0),
      orderHariIni: todayOrders.length,
      omzetHariIni: todayOrders.reduce((sum, order) => sum + Number(order.Total || 0), 0),
      omzetTotal: orders.reduce((sum, order) => sum + Number(order.Total || 0), 0),
      stokMenipis: products.filter(product => Number(product.Aktif) !== 0 && Number(product.Stok || 0) <= 5).length
    }
  };
}

function findProduct_(sku) {
  const sheet = SpreadsheetApp.getActive().getSheetByName('MASTER_PRODUK');
  const values = sheet.getDataRange().getValues();
  for (let row = 1; row < values.length; row++) {
    if (String(values[row][0]) === String(sku)) return { sheet, row: row + 1, values: values[row] };
  }
  throw new Error('Produk tidak ditemukan: ' + sku);
}

function addProduct_(p) {
  if (!p.SKU || !p['Nama Produk']) throw new Error('SKU dan Nama Produk wajib diisi.');
  const sheet = SpreadsheetApp.getActive().getSheetByName('MASTER_PRODUK');
  if (rows_('MASTER_PRODUK').some(item => String(item.SKU) === String(p.SKU))) throw new Error('SKU sudah digunakan.');
  sheet.appendRow([p.SKU,p['Nama Produk'],p.Kategori || '',Number(p['Harga Modal'] || 0),Number(p['Harga Jual'] || 0),Number(p.Stok || 0),Number(p.Aktif ?? 1),new Date(),p.Brand || 'Tanpa Brand']);
  if (Number(p.Stok || 0) > 0) appendMutation_(p.SKU, 'STOK_AWAL', Number(p.Stok), p.SKU, 'Tambah produk');
  return { success: true, SKU: p.SKU };
}

function updateProduct_(p) {
  const found = findProduct_(p.SKU);
  const old = found.values;
  found.sheet.getRange(found.row, 1, 1, 9).setValues([[
    p.SKU, p['Nama Produk'] ?? old[1], p.Kategori ?? old[2], Number(p['Harga Modal'] ?? old[3]),
    Number(p['Harga Jual'] ?? old[4]), Number(p.Stok ?? old[5]), Number(p.Aktif ?? old[6]), new Date(), p.Brand ?? old[8] ?? 'Tanpa Brand'
  ]]);
  return { success: true, SKU: p.SKU };
}

function nonaktifProduct_(p) {
  const found = findProduct_(p.SKU);
  found.sheet.getRange(found.row, 7).setValue(0);
  found.sheet.getRange(found.row, 8).setValue(new Date());
  return { success: true, SKU: p.SKU };
}

function addStock_(p) {
  const qty = Number(p.Qty || 0);
  if (qty <= 0) throw new Error('Qty stok harus lebih dari 0.');
  const found = findProduct_(p.SKU);
  found.sheet.getRange(found.row, 6).setValue(Number(found.values[5] || 0) + qty);
  found.sheet.getRange(found.row, 8).setValue(new Date());
  appendMutation_(p.SKU, 'MASUK', qty, p.Referensi || '', p.Keterangan || 'Tambah stok');
  return { success: true, SKU: p.SKU, stock: Number(found.values[5] || 0) + qty };
}

function addBonus_(p) {
  const qty = Number(p.Qty || 0);
  if (qty <= 0) throw new Error('Qty bonus harus lebih dari 0.');
  const found = findProduct_(p.SKU);
  const stock = Number(found.values[5] || 0);
  if (stock < qty) throw new Error('Stok tidak cukup untuk bonus: ' + p.SKU);
  found.sheet.getRange(found.row, 6).setValue(stock - qty);
  found.sheet.getRange(found.row, 8).setValue(new Date());
  appendMutation_(p.SKU, 'KELUAR', -qty, p.Referensi || '', 'Bonus: ' + (p.Keterangan || 'Bonus pelanggan'));
  return { success: true, SKU: p.SKU, stock: stock - qty, bonusQty: qty };
}

function addBonuses_(payload) {
  const items = payload.items || [];
  if (!items.length) throw new Error('Produk bonus kosong.');
  items.forEach(item => addBonus_({
    SKU: item.SKU,
    Qty: item.Qty,
    Referensi: payload.Referensi || '',
    Keterangan: payload.Keterangan || 'Bonus pelanggan'
  }));
  return { success: true, itemCount: items.length };
}

function saveOrder_(payload) {
  if (!payload.items || !payload.items.length) throw new Error('Item order kosong.');
  const ss = SpreadsheetApp.getActive();
  const orderId = 'ORD-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss');
  let subtotal = 0;
  const itemRows = payload.items.map(item => {
    const found = findProduct_(item.SKU);
    const qty = Number(item.Qty || 0);
    const stock = Number(found.values[5] || 0);
    if (qty <= 0) throw new Error('Qty tidak valid: ' + item.SKU);
    if (stock < qty) throw new Error('Stok tidak cukup: ' + item.SKU);
    const price = Number(found.values[4] || 0);
    subtotal += price * qty;
    found.sheet.getRange(found.row, 6).setValue(stock - qty);
    found.sheet.getRange(found.row, 8).setValue(new Date());
    appendMutation_(item.SKU, 'KELUAR', -qty, orderId, 'Penjualan');
    return [orderId,item.SKU,found.values[1],price,qty,price * qty];
  });
  const total = subtotal + Number(payload.Ongkir || 0) - Number(payload.Diskon || 0);
  ss.getSheetByName('ORDER').appendRow([orderId,new Date(),payload['Nama Customer'] || '',payload['No WhatsApp'] || '',payload.Alamat || '',payload.Ekspedisi || '',Number(payload.Ongkir || 0),Number(payload.Diskon || 0),subtotal,total,payload['Status Pembayaran'] || 'Belum Bayar',payload['Status Order'] || 'Diproses',payload.Catatan || '']);
  ss.getSheetByName('ORDER_ITEM').getRange(ss.getSheetByName('ORDER_ITEM').getLastRow() + 1, 1, itemRows.length, 6).setValues(itemRows);
  upsertCustomer_(payload);
  return { success: true, orderId, totalOrder: total };
}

function getOrders_() {
  const items = rows_('ORDER_ITEM');
  return rows_('ORDER').map(order => ({ ...order, items: items.filter(item => item['Order ID'] === order['Order ID']).map(item => ({ sku: item.SKU, name: item['Nama Produk'], price: item.Harga, qty: item.Qty, total: item.Total })) }));
}

function deleteOrders_(orderIds) {
  const ids = [...new Set((orderIds || []).map(String).filter(Boolean))];
  if (!ids.length) throw new Error('Order ID wajib diisi.');
  const ss = SpreadsheetApp.getActive();
  const deletedOrders = deleteRowsByIds_(ss.getSheetByName('ORDER'), ids);
  deleteRowsByIds_(ss.getSheetByName('ORDER_ITEM'), ids);
  if (!deletedOrders) throw new Error('Order tidak ditemukan.');
  return { success: true, deleted: deletedOrders, orderIds: ids };
}

function deleteRowsByIds_(sheet, ids) {
  if (!sheet || sheet.getLastRow() < 2) return 0;
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  let deleted = 0;
  for (let index = values.length - 1; index >= 0; index--) {
    if (ids.includes(String(values[index][0]))) {
      sheet.deleteRow(index + 2);
      deleted++;
    }
  }
  return deleted;
}

function appendMutation_(sku, type, qty, reference, note) {
  SpreadsheetApp.getActive().getSheetByName('MUTASI_STOK').appendRow([new Date(),sku,type,qty,reference,note]);
}

function upsertCustomer_(p) {
  const sheet = SpreadsheetApp.getActive().getSheetByName('PELANGGAN');
  const values = sheet.getDataRange().getValues();
  for (let row = 1; row < values.length; row++) {
    if (String(values[row][0]) === String(p['No WhatsApp'])) {
      sheet.getRange(row + 1, 1, 1, 4).setValues([[p['No WhatsApp'],p['Nama Customer'],p.Alamat,new Date()]]);
      return;
    }
  }
  sheet.appendRow([p['No WhatsApp'],p['Nama Customer'],p.Alamat,new Date()]);
}

function json_(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function getDb() {
  const props = PropertiesService.getScriptProperties();
  let dbId = props.getProperty('DB_ID');
  let ss;

  if (!dbId) {
    ss = SpreadsheetApp.create('EduPOS_DB');
    props.setProperty('DB_ID', ss.getId());
    setupDatabase(ss);
  } else {
    try {
      ss = SpreadsheetApp.openById(dbId);
    } catch(e) {
      ss = SpreadsheetApp.create('EduPOS_DB');
      props.setProperty('DB_ID', ss.getId());
      setupDatabase(ss);
    }
  }
  return ss;
}

function setupDatabase(ss) {
  const sheets = ['Products', 'Transactions', 'Logs', 'Settings'];
  sheets.forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      if (name === 'Products') {
        sheet.appendRow(['id', 'name', 'price', 'stock', 'category', 'image_url']);
      } else if (name === 'Transactions') {
        sheet.appendRow(['id', 'date', 'total_amount', 'items', 'payment_method']);
      } else if (name === 'Logs') {
        sheet.appendRow(['id', 'date', 'action', 'details']);
      } else if (name === 'Settings') {
        sheet.appendRow(['key', 'value']);
        sheet.appendRow(['storeName', 'EDUPOS STORE']);
        sheet.appendRow(['storeAddress', 'Jl. Teknologi No. 123, Jakarta']);
        sheet.appendRow(['storePhone', '0812-3456-7890']);
        sheet.appendRow(['taxRate', '10']);
      }
    }
  });

  // Remove default 'Sheet1' if it exists
  const sheet1 = ss.getSheetByName('Sheet1');
  if (sheet1 && ss.getSheets().length > 1) {
    ss.deleteSheet(sheet1);
  }
}

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('EduPOS - Smart Cashier')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getProducts() {
  const sheet = getDb().getSheetByName('Products');
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    let obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}

function saveProduct(product) {
  const ss = getDb();
  const sheet = ss.getSheetByName('Products');

  if (product.id) {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == product.id) {
        sheet.getRange(i + 1, 2, 1, 5).setValues([[product.name, product.price, product.stock, product.category, product.image_url || '']]);
        logAction('Update Product', `Updated product: ${product.name}`);
        return { success: true };
      }
    }
  }

  const newId = 'PRD' + new Date().getTime();
  sheet.appendRow([newId, product.name, product.price, product.stock, product.category, product.image_url || '']);
  logAction('Add Product', `Added new product: ${product.name}`);
  return { success: true };
}

function deleteProduct(id) {
  const sheet = getDb().getSheetByName('Products');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      const name = data[i][1];
      sheet.deleteRow(i + 1);
      logAction('Delete Product', `Deleted product: ${name}`);
      return { success: true };
    }
  }
  return { success: false, error: 'Product not found' };
}

function getTransactions() {
  const sheet = getDb().getSheetByName('Transactions');
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    let obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  }).reverse();
}

function processCheckout(cart, total, paymentMethod) {
  const ss = getDb();
  const txSheet = ss.getSheetByName('Transactions');
  const prdSheet = ss.getSheetByName('Products');

  const txId = 'TRX' + new Date().getTime();
  const date = new Date().toISOString();

  const prdData = prdSheet.getDataRange().getValues();
  cart.forEach(item => {
    for (let i = 1; i < prdData.length; i++) {
      if (prdData[i][0] == item.id) {
        const currentStock = prdData[i][3];
        prdSheet.getRange(i + 1, 4).setValue(currentStock - item.qty);
        break;
      }
    }
  });

  txSheet.appendRow([txId, date, total, JSON.stringify(cart), paymentMethod]);
  logAction('Checkout', `Processed transaction ${txId} for Rp ${total}`);

  return { success: true, transactionId: txId, date: date };
}

function getLogs() {
  const sheet = getDb().getSheetByName('Logs');
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return [];
  const headers = data[0];
  return data.slice(1).map(row => {
    let obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  }).reverse();
}

function logAction(action, details) {
  const sheet = getDb().getSheetByName('Logs');
  const id = 'LOG' + new Date().getTime();
  sheet.appendRow([id, new Date().toISOString(), action, details]);
}

function getSettings() {
  const sheet = getDb().getSheetByName('Settings');
  const data = sheet.getDataRange().getValues();
  let settings = {};
  for (let i = 1; i < data.length; i++) {
    settings[data[i][0]] = data[i][1];
  }
  return settings;
}

function saveSettings(settings) {
  const sheet = getDb().getSheetByName('Settings');
  const data = sheet.getDataRange().getValues();

  const keys = Object.keys(settings);
  for (let k of keys) {
    let found = false;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === k) {
        sheet.getRange(i + 1, 2).setValue(settings[k]);
        found = true;
        break;
      }
    }
    if (!found) {
      sheet.appendRow([k, settings[k]]);
    }
  }
  logAction('Update Settings', 'Store settings updated');
  return { success: true };
}

function getDashboardSummary() {
  const ss = getDb();
  const txSheet = ss.getSheetByName('Transactions');
  const prdSheet = ss.getSheetByName('Products');

  const txData = txSheet.getDataRange().getValues();
  const prdData = prdSheet.getDataRange().getValues();

  let totalSales = 0;
  let txCount = 0;

  // Chart Data preparation (Last 7 days)
  let salesByDate = {};
  for(let i=0; i<7; i++) {
    let d = new Date();
    d.setDate(d.getDate() - i);
    let dateStr = Utilities.formatDate(d, "Asia/Jakarta", "yyyy-MM-dd");
    salesByDate[dateStr] = 0;
  }

  if (txData.length > 1) {
    txCount = txData.length - 1;
    for (let i = 1; i < txData.length; i++) {
      const amount = Number(txData[i][2]) || 0;
      totalSales += amount;

      const txDate = new Date(txData[i][1]);
      const dateStr = Utilities.formatDate(txDate, "Asia/Jakarta", "yyyy-MM-dd");
      if (salesByDate[dateStr] !== undefined) {
        salesByDate[dateStr] += amount;
      }
    }
  }

  let chartLabels = Object.keys(salesByDate).reverse();
  let chartData = chartLabels.map(d => salesByDate[d]);

  let totalProducts = prdData.length > 1 ? prdData.length - 1 : 0;

  let recentTx = [];
  if (txData.length > 1) {
    recentTx = txData.slice(-5).map(row => ({
      id: row[0],
      date: row[1],
      total: row[2]
    })).reverse();
  }

  return { totalSales, txCount, totalProducts, recentTx, chartLabels, chartData };
}

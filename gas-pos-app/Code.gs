function getDb() {
  // Karena script ini dijalankan dari dalam Spreadsheet (Bound Script)
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  setupDatabase(ss);
  return ss;
}

function setupDatabase(ss) {
  const sheets = ['Products', 'Transactions', 'Logs', 'Settings'];
  sheets.forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) {
      sheet = ss.insertSheet(name);
      if (name === 'Products') {
        // Updated schema
        sheet.appendRow(['id', 'barcode', 'name', 'cost_price', 'price', 'wholesale_price', 'stock', 'category', 'type', 'image_url']);
      } else if (name === 'Transactions') {
        // Updated schema
        sheet.appendRow(['id', 'date', 'subtotal', 'tax', 'discount', 'total_amount', 'paid_amount', 'change', 'items', 'payment_method', 'profit']);
      } else if (name === 'Logs') {
        sheet.appendRow(['id', 'date', 'action', 'details']);
      } else if (name === 'Settings') {
        sheet.appendRow(['key', 'value']);
        sheet.appendRow(['storeName', 'EDUPOS STORE']);
        sheet.appendRow(['storeAddress', 'Jl. Teknologi No. 123, Jakarta']);
        sheet.appendRow(['storePhone', '0812-3456-7890']);
        sheet.appendRow(['taxRate', '10']);
        sheet.appendRow(['storeLogo', '']);
      }
    } else {
      // Ensure missing columns are added without breaking existing rows
      const lastCol = sheet.getLastColumn();
      if (lastCol > 0) {
        const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

        if (name === 'Products') {
          const required = ['id', 'barcode', 'name', 'cost_price', 'price', 'wholesale_price', 'stock', 'category', 'type', 'image_url'];
          required.forEach(col => {
            if (!headers.includes(col)) {
              sheet.getRange(1, sheet.getLastColumn() + 1).setValue(col);
            }
          });
        } else if (name === 'Transactions') {
          const required = ['id', 'date', 'subtotal', 'tax', 'discount', 'total_amount', 'paid_amount', 'change', 'items', 'payment_method', 'profit'];
          required.forEach(col => {
            if (!headers.includes(col)) {
              sheet.getRange(1, sheet.getLastColumn() + 1).setValue(col);
            }
          });
        } else if (name === 'Settings') {
          // Add default settings if not exists
          const data = sheet.getDataRange().getValues();
          const existingKeys = data.map(r => r[0]);
          if (!existingKeys.includes('storeLogo')) sheet.appendRow(['storeLogo', '']);
        }
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
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const rowData = headers.map(h => {
    if (h === 'id') return product.id || ('PRD' + new Date().getTime());
    return product[h] !== undefined ? product[h] : '';
  });

  if (product.id) {
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] == product.id) {
        sheet.getRange(i + 1, 1, 1, headers.length).setValues([rowData]);
        logAction('Update Product', `Updated product: ${product.name}`);
        return { success: true };
      }
    }
  }

  sheet.appendRow(rowData);
  logAction('Add Product', `Added new product: ${product.name}`);
  return { success: true };
}

function deleteProduct(id) {
  const sheet = getDb().getSheetByName('Products');
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      const nameIdx = data[0].indexOf('name');
      const name = data[i][nameIdx];
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

function processCheckout(txData) {
  const ss = getDb();
  const txSheet = ss.getSheetByName('Transactions');
  const prdSheet = ss.getSheetByName('Products');

  const txId = 'TRX' + new Date().getTime();
  const date = new Date().toISOString();

  const prdData = prdSheet.getDataRange().getValues();
  const prdHeaders = prdData[0];
  const idIdx = prdHeaders.indexOf('id');
  const stockIdx = prdHeaders.indexOf('stock');
  const costIdx = prdHeaders.indexOf('cost_price');

  let totalProfit = 0;

  txData.cart.forEach(item => {
    for (let i = 1; i < prdData.length; i++) {
      if (prdData[i][idIdx] == item.id) {
        // Update stock
        const currentStock = Number(prdData[i][stockIdx]) || 0;
        prdSheet.getRange(i + 1, stockIdx + 1).setValue(currentStock - item.qty);

        // Calculate profit
        const costPrice = Number(prdData[i][costIdx]) || 0;
        const sellingPrice = Number(item.appliedPrice || item.price) || 0;
        totalProfit += (sellingPrice - costPrice) * item.qty;
        break;
      }
    }
  });

  // Deduct transaction discount from total profit
  totalProfit -= (Number(txData.discount) || 0);

  const txHeaders = txSheet.getRange(1, 1, 1, txSheet.getLastColumn()).getValues()[0];
  const rowData = txHeaders.map(h => {
    switch(h) {
      case 'id': return txId;
      case 'date': return date;
      case 'subtotal': return txData.subtotal;
      case 'tax': return txData.tax;
      case 'discount': return txData.discount;
      case 'total_amount': return txData.total;
      case 'paid_amount': return txData.paid_amount;
      case 'change': return txData.change;
      case 'items': return JSON.stringify(txData.cart);
      case 'payment_method': return txData.paymentMethod;
      case 'profit': return totalProfit;
      default: return '';
    }
  });

  txSheet.appendRow(rowData);
  logAction('Checkout', `Processed transaction ${txId} for Rp ${txData.total}`);

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

  // Best selling logic & Low stock
  let productSalesCount = {};
  let lowStockProducts = [];

  const prdHeaders = prdData[0] || [];
  const stockIdx = prdHeaders.indexOf('stock');
  const nameIdx = prdHeaders.indexOf('name');
  const idIdx = prdHeaders.indexOf('id');

  if (prdData.length > 1) {
    for (let i = 1; i < prdData.length; i++) {
      const st = Number(prdData[i][stockIdx]) || 0;
      if (st <= 5) {
        lowStockProducts.push({
          id: prdData[i][idIdx],
          name: prdData[i][nameIdx],
          stock: st
        });
      }
    }
  }

  let recentTx = [];
  if (txData.length > 1) {
    const txHeaders = txData[0];
    const itemsIdx = txHeaders.indexOf('items');

    // Calculate best selling from all tx
    for (let i = 1; i < txData.length; i++) {
      try {
        let items = JSON.parse(txData[i][itemsIdx]);
        items.forEach(it => {
          if (!productSalesCount[it.name]) productSalesCount[it.name] = 0;
          productSalesCount[it.name] += it.qty;
        });
      } catch(e) {}
    }

    recentTx = txData.slice(-5).map(row => {
      let obj = {};
      txHeaders.forEach((h, i) => obj[h] = row[i]);
      return obj;
    }).reverse();
  }

  let topProducts = Object.keys(productSalesCount)
    .map(name => ({ name, qty: productSalesCount[name] }))
    .sort((a, b) => b.qty - a.qty)
    .slice(0, 5);

  return { totalSales, txCount, totalProducts, recentTx, chartLabels, chartData, topProducts, lowStockProducts };
}

function getReports(period) {
  // period: 'daily', 'weekly', 'monthly'
  const ss = getDb();
  const txSheet = ss.getSheetByName('Transactions');
  const txData = txSheet.getDataRange().getValues();

  let sales = 0;
  let profit = 0;
  let count = 0;

  if (txData.length <= 1) return { sales, profit, count };

  const headers = txData[0];
  const dateIdx = headers.indexOf('date');
  const totalIdx = headers.indexOf('total_amount');
  const profitIdx = headers.indexOf('profit');

  const now = new Date();

  for (let i = 1; i < txData.length; i++) {
    const txDate = new Date(txData[i][dateIdx]);
    let include = false;

    if (period === 'daily') {
      if (txDate.toDateString() === now.toDateString()) include = true;
    } else if (period === 'weekly') {
      const diffTime = Math.abs(now - txDate);
      const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
      if (diffDays <= 7) include = true;
    } else if (period === 'monthly') {
      if (txDate.getMonth() === now.getMonth() && txDate.getFullYear() === now.getFullYear()) include = true;
    }

    if (include) {
      sales += Number(txData[i][totalIdx]) || 0;
      profit += Number(txData[i][profitIdx]) || 0;
      count++;
    }
  }

  return { sales, profit, count };
}

function getBackupData() {
  const ss = getDb();
  const sheets = ['Products', 'Transactions', 'Logs', 'Settings'];
  let backup = {};

  sheets.forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (sheet) {
      backup[name] = sheet.getDataRange().getValues();
    }
  });

  return JSON.stringify(backup);
}

function restoreData(jsonData) {
  try {
    const data = JSON.parse(jsonData);
    const ss = getDb();

    Object.keys(data).forEach(name => {
      let sheet = ss.getSheetByName(name);
      if (sheet) {
        sheet.clear();
        sheet.getRange(1, 1, data[name].length, data[name][0].length).setValues(data[name]);
      }
    });

    logAction('System Restore', 'Database restored from backup');
    return { success: true };
  } catch(e) {
    return { success: false, error: e.toString() };
  }
}

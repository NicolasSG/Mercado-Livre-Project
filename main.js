// Configuration constants
const CONFIG = {
  API: {
    URL: 'https://api.mercadolibre.com',
    TOKEN_ENDPOINT: '/oauth/token',
    ORDERS_ENDPOINT: '/orders',
    SELLER_ID: 'YOUR SELLER ID'
  },
  SHEETS: {
    CLIENTS: "YOUR SHEET ID",
    ALUGUEL: "YOUR SHEET ID",
    FINANCEIRO: "YOUR SHEET ID",
    CLIENT_SHEET_NAME: "Clientes",
    ENTRADAS_SHEET_NAME: "Entradas"
  },
  AUTH: {
    CLIENT_ID: 'YOUR ID',
    CLIENT_SECRET: 'YOUR ID',
    GRANT_TYPE: 'refresh_token',
    REFRESH_TOKEN: 'YOUR TOKEN',
    SCOPE: 'offline_access read write'
  }
};

// Helper functions
const Utils = {
  objectToQueryParams: (obj) => {
    return Object.entries(obj)
      .map(([k, v]) => `${encodeURIComponent(k)}=${encodeURIComponent(v)}`)
      .join('&');
  },
  
  lastNumber: (string) => {
    const n = string.split(/[\s,]+/);  
    return n[n.length - 1];
  },
  
  formatDate: (date, daysToAdd = 0) => {
    const newDate = new Date(date);
    newDate.setDate(newDate.getDate() + daysToAdd);
    return Utilities.formatDate(newDate, "GMT", "dd/MM/yyyy");
  }
};

// API Service
const MercadoLivreService = {
  getAccessToken: () => {
    const headers = {
      Accept: 'application/json',
      'Content-Type': 'application/x-www-form-urlencoded',
    };

    const requestOptions = {
      method: 'POST',
      headers,
      payload: Utils.objectToQueryParams(CONFIG.AUTH),
      followRedirects: true,
    };

    const response = UrlFetchApp.fetch(CONFIG.API.URL + CONFIG.API.TOKEN_ENDPOINT, requestOptions);
    return JSON.parse(response.getContentText()).access_token;
  },

  getAuthHeader: (token) => ({
    method: "GET",
    headers: {
      "Authorization": `Bearer ${token}`,
      "cache-control": "no-cache"
    }
  }),

  getSells: (token) => {
    const options = MercadoLivreService.getAuthHeader(token);
    const url = `${CONFIG.API.URL}${CONFIG.API.ORDERS_ENDPOINT}/search?seller=${CONFIG.API.SELLER_ID}&sort=date_desc`;
    return JSON.parse(UrlFetchApp.fetch(url, options).getContentText());
  },

  getOrder: (token, orderId) => {
    const options = MercadoLivreService.getAuthHeader(token);
    const url = `${CONFIG.API.URL}${CONFIG.API.ORDERS_ENDPOINT}/${orderId}`;
    return JSON.parse(UrlFetchApp.fetch(url, options).getContentText());
  }
};

// Spreadsheet Service
const SpreadsheetService = {
  getOrderInfo: (orderParser) => {
    const clientSheet = SpreadsheetApp.openById(CONFIG.SHEETS.CLIENTS).getSheets()[0];
    const lastRow = clientSheet.getLastRow();
    const lastRowText = clientSheet.getRange(lastRow, 1).getValue();  
    const clientNumber = parseInt(Utils.lastNumber(lastRowText), 10);

    const buyer = orderParser.buyer;
    const name = `${buyer.first_name} ${buyer.last_name} - ${clientNumber + 1}`;
    const nickname = buyer.nickname;
    const phone = `${buyer.phone.area_code} ${buyer.phone.number}`;
    const totalPayment = orderParser.paid_amount;
    const initialDate = Utils.formatDate(orderParser.date_created, 1);
    const finalDate = Utils.formatDate(orderParser.date_created, 7);
    const orderId = orderParser.id;

    return [nickname, "7 dias", totalPayment, initialDate, finalDate, name, phone, "", orderId];
  },

  checkIDinSheets: (allSheets, sellParser) => {
    const ordersList = [];
    
    for (let i = 0; i > -1; i++) {
      const checkLastID = sellParser.results[i].payments[0].order_id;
      
      for (let j = 0; j < allSheets.length; j++) {
        const sheet = allSheets[j];
        const found = sheet.createTextFinder(checkLastID).findAll()[0];     
        if (found) {
          break;
        }
      }
      ordersList.push(checkLastID);
    }  
    return ordersList;
  },

  matchSheetName: (allSheets, productName) => {
    for (let i = 0; i < allSheets.length; i++) {
      const name = allSheets[i].getSheetName();
      const nameSliced = name.slice(0, -3);

      if (productName.indexOf(nameSliced) !== -1) {
        if (name.includes("#1") && productName.includes("#1")) {
          return name;
        }
        if (name.includes("#2") && productName.includes("#2")) {
          return name;
        }
        if (productName.includes("Pc Digital")) {
          return name;      
        }
      }
    }
  },

  saveData: (sheetName, rowInfo) => {
    const sheet = SpreadsheetApp.openById(CONFIG.SHEETS.ALUGUEL).getSheetByName(sheetName);  
    const appendedRow = sheet.appendRow(rowInfo);
    const lastRow = appendedRow.getLastRow();
    appendedRow.getRange(lastRow, 8).setBackground("#6aa84f");  
  },

  createClientSheet: (orderInfo, productName) => {
    const clientNumber = parseInt(Utils.lastNumber(orderInfo[5]));
    const sheet = SpreadsheetApp.openById(CONFIG.SHEETS.CLIENTS);
    const mainSheet = sheet.getSheetByName(CONFIG.SHEETS.CLIENT_SHEET_NAME);
    
    const newInfo = [orderInfo[5], orderInfo[6], orderInfo[0]];
    const tabInfo1 = ["Histórico", "Dia", "Pagou como?", "Valor", "Jogo", "ID da transação PIX", "ID da venda ML"];
    const tabInfo2 = ["Comprou locação", orderInfo[3], "Mercado Livre", orderInfo[2], productName, "", orderInfo[8]];

    const newSheet = sheet.insertSheet(orderInfo[5], clientNumber);
    mainSheet.appendRow(newInfo);
    newSheet.appendRow(tabInfo1);
    newSheet.appendRow(tabInfo2);  
    
    // Formatting
    newSheet.getRange("1:1")
      .setBackground("#cccccc")
      .setFontWeight("bold")
      .setFontSize(13)
      .setHorizontalAlignment("center");
    newSheet.getRange("2:1").setHorizontalAlignment("center");  
    newSheet.getRange("D2:D").setNumberFormat("R$ ##,##");

    // Copy column widths from template
    const templateSheet = sheet.getSheets()[1];
    for (let i = 1; i <= templateSheet.getLastColumn(); i++) {
      newSheet.setColumnWidth(i, templateSheet.getColumnWidth(i));
    }
  },

  updateFinanceSheet: (orderInfo, productName, orderParser) => {
    const sheet = SpreadsheetApp.openById(CONFIG.SHEETS.FINANCEIRO)
      .getSheetByName(CONFIG.SHEETS.ENTRADAS_SHEET_NAME);
    const lastRow = sheet.getLastRow();
    
    const positionInArray = new Date(orderParser.date_created).getMonth();  
    const newRow = [orderInfo[3], "ML", productName];
    
    for (let i = 0; i < positionInArray; i++) {
      newRow.push("");
    }
    
    const paymentWithDiscount = parseInt(orderInfo[2]) - orderParser.order_items[0].sale_fee;  
    newRow.push(paymentWithDiscount);
    
    sheet.insertRowBefore(lastRow);
    sheet.getRange(lastRow, 1, 1, newRow.length).setValues([newRow]);
    sheet.getRange(lastRow, 1, 1, 18).setBackground("white");

    // Update summary formulas
    for (let col = 4; col <= 15; col++) {
      const columnLetter = String.fromCharCode(64 + col);
      sheet.getRange(lastRow, col).setFormula(`=SUM(${columnLetter}1:${columnLetter}${lastRow-1})`);
    }
    
    sheet.getRange(lastRow, 16).setFormula(`=IFERROR(Q${lastRow}/COUNTA(D${lastRow}:O${lastRow}),0)`);
    sheet.getRange(lastRow, 17).setFormula(`=SUM(D${lastRow}:O${lastRow})`);
  }
};

// Main function to process orders
function processOrders() {
  const accessToken = MercadoLivreService.getAccessToken();
  const sellParser = MercadoLivreService.getSells(accessToken);
  
  const aluguelSpreadsheet = SpreadsheetApp.openById(CONFIG.SHEETS.ALUGUEL);
  const allSheets = aluguelSpreadsheet.getSheets();
  const orderIDs = SpreadsheetService.checkIDinSheets(allSheets, sellParser);

  orderIDs.forEach(orderId => {
    const order = MercadoLivreService.getOrder(accessToken, orderId);
    const productName = order.order_items[0].item.title;
    const orderInfo = SpreadsheetService.getOrderInfo(order);
    const sheetName = SpreadsheetService.matchSheetName(allSheets, productName);
    
    SpreadsheetService.saveData(sheetName, orderInfo);
    SpreadsheetService.createClientSheet(orderInfo, productName);
    SpreadsheetService.updateFinanceSheet(orderInfo, productName, order);
  });
}

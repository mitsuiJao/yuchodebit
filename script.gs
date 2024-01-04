const address = "yuchodebit@jp-bank.japanpost.jp";
const ssId = "" //spreadsheet id
const today = new Date();
const newsheet = true; //false:auto true:create


function main() {
  let threads = getMailThreads(address);
  let contents = getThreadContents(threads);
  let ss = getSpreadsheet(ssId);
  let sheet = getSheet(ss);
  write(sheet, contents, "A2:D2");
  summary(ss);
  console.log("done");
}


function summary(ss) {
  ss = getSpreadsheet(ssId);
  let sheet = ss.getActiveSheet();
  let byStoreSum = {};
  let byDateSum = {};
  let Dresult = [];
  let Sresult = [];
  let row = 2; 
  while (true) { //検索対象が空で無かったら
    let storeRange = sheet.getRange("C" + String(row));
    let dateRange = sheet.getRange("B" + String(row));
    //console.log("pass1");
    if (storeRange.isBlank()) {
      break;
    }

    let storeStore = String(storeRange.getValue());
    let storeExpence = storeRange.offset(0, 1).getValue();
    let dateDrift = dateRange.getValue();
    let dateStore = `${dateDrift.getFullYear()}/${dateDrift.getMonth() + 1}/${dateDrift.getDate()}`
    let dateExpence = dateRange.offset(0, 2).getValue();

    if (!Object.keys(byStoreSum).includes(storeStore)) { //検索対象の店が配列に含まれてなかったら
      byStoreSum[storeStore] = storeExpence;
    } else {
      byStoreSum[storeStore] += storeExpence;
    }

    if (!Object.keys(byDateSum).includes(dateStore)) {
      byDateSum[dateStore] = dateExpence;
    } else {
      byDateSum[dateStore] += dateExpence;
    }

    //byDateSumArr.push(byDateSum);
    //byStoreSumArr.push(byStoreSum);
    row++;
  }

  //console.log(byDateSum, byStoreSum)
  for (let date in byDateSum){
    let val = byDateSum[date];
    let pushObj = {
      "time": date,
      "expence": val
    }
    Dresult.push(pushObj);
  }
  for (let store in byStoreSum){
    let val = byStoreSum[store];
    let pushObj = {
      "shop": store,
      "expence": val
    }
    Sresult.push(pushObj);
  }

  sheet.getRange("H:K").clear();
  write(sheet, Sresult, "J1:K1");
  write(sheet, Dresult, "H1:I1");
  return Dresult, Sresult;
}


function getMailThreads(address) {
  let yesterday = new Date();
  yesterday.setDate(yesterday.getDate()-1);
  let datey = yesterday.getFullYear();
  let datem = yesterday.getMonth()+1;
  let dated = yesterday.getDate();
  let date = datey + "/" + datem + "/" + dated; //dated
  let straddress = address+" after:"+date;
  if (newsheet){
    let tmp = datey + "/" + datem + "/" + "1";
    straddress = address+" after:"+tmp
  } 
  console.log(straddress);
  let threads = GmailApp.search(straddress);
  console.log("GetMail...done");
  return threads;
}


function getThreadContents(threads) {
  let result = []
  threads.forEach(function (thread) {
    let messages = thread.getMessages();

    messages.forEach(function (message) {
      let body = message.getPlainBody();
      let when = find(body, 1);
      let shop = find(body, 2);
      let expense = find(body, 3);
      let id = createId(when);
      let temp = {
        "id": id,
        "date": when,
        "shop": shop,
        "expense": expense
      }
      if (!isNaN(temp.expense)) {
        result.push(temp);
      }
    });
  });
  return result;
}


function getSpreadsheet() { //spreadsheetApp作成
  let spreadsheet = SpreadsheetApp.openById(ssId);
  return spreadsheet;
}


function getSheet(spreadsheet) {  //月一で初期化、作成　またはアクティブシートを取得
  let flg;
  if (newsheet){
    flg = 1
  } else {
    flg = today.getDate();
  }
  if (flg == 1) { //today.getDate() == 1
    let createSheetName =
      String(today.getFullYear()) + "/" + String(today.getMonth() + 1);

    spreadsheet.insertSheet(createSheetName, 0);

    let sheet = spreadsheet.getActiveSheet();
    sheet.getRange("A1:D1").setValues([["id", "when", "shop", "expence"]]);

    write(sheet, "total", "F1");
    sheet.getRange("F2").setFormula("sum(D:D)");
    write(sheet, "mostStore", "F5");
    sheet.getRange("F6").setFormula("=INDEX(J:J,MATCH(F7,K:K,0),1)");
    sheet.getRange("F7").setFormula("MAX(K:K)");
    write(sheet, "mostDate", "F10");
    sheet.getRange("F11").setFormula("=INDEX(H:H,MATCH(F12,I:I,0),1)");
    sheet.getRange("F12").setFormula("MAX(I:I)");

    let tmp = sheet.getRange("F1");
    tmp.setFontWeight("bold");
    tmp.setFontSize(12);
    tmp = sheet.getRange("F5");
    tmp.setFontWeight("bold");
    tmp.setFontSize(12);
    tmp = sheet.getRange("F10");
    tmp.setFontWeight("bold");
    tmp.setFontSize(12);
    tmp = sheet.getRange("F:F")
    tmp.setHorizontalAlignment("center");


    sheet.getRange("A1:D2").createFilter();
    sheet.setColumnWidth(1, 125);
    sheet.setColumnWidth(3, 280);
    sheet.setColumnWidth(6, 280);
    sheet.setColumnWidth(10, 280);

  }
  let sheet = spreadsheet.getActiveSheet();
  return sheet;
}


function write(sheet, object, range) {
  if (typeof (object) == "object") {
    object.forEach(function(obj){
      let content = Object.values(obj);
      sheet.getRange(range).insertCells(SpreadsheetApp.Dimension.ROWS);
      let wirteRange = sheet.getRange(range);
      wirteRange.setValues([content]);
    });
  } else {
    sheet.getRange(range).setValue(object);
  }
}


function find(str, mode) {//日時:1 店舗:2 金額:3
  const modeDic = {
    1: "利用日時",
    2: "利用店舗",
    3: "利用金額"
  }
  let sIndex = str.indexOf(modeDic[mode]);
  let eIndex = str.indexOf("\r", sIndex + 1);
  let matched = str.slice(sIndex + 6, eIndex);
  if (mode == 1) {
    matched = new Date(matched);
  } else if (mode == 3) {
    matched = matched.slice(0, matched.length - 1);
    matched = matched.replace(",", "");
    matched = Number(matched)
  }
  return matched;
}


function createId(date) {
  let yyyy = String(date.getFullYear());
  let mm = String(date.getMonth() + 1).padStart(2, "0");
  let dd = String(date.getDate()).padStart(2, "0");
  let hh = String(date.getHours()).padStart(2, "0");
  let mi = String(date.getMinutes()).padStart(2, "0");
  let ss = String(date.getSeconds()).padStart(2, "0");
  return yyyy + mm + dd + hh + mi + ss;
}

{
  /* 
  mail
  [
    {
      id:id
      date:date
      store:store
      expence:expence
    },
    {
      id:id
      date:date
      store:store
      expence:expence
    },,,
  ]

  summary
  [
    {
      time:time
      expence:expence
    },
    {
      time:time
      expence:expence
    },
    {
      time:time
      expence:expence
    },,,
  ]
  */
}

/* その日だけを取得して記述、summaryは一度全部消します！*/
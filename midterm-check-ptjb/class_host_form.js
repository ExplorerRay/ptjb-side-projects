function createNewForm2(){
  var ui = SpreadsheetApp.getUi();
  let yr = ui.prompt('請輸入年度').getResponseText();
  // 下學期
  let new_form_2 = FormApp.create(yr + "-2 課程查核表單"); 
  let new_form_2_id = new_form_2.getId();

  let destinationFolderUrl = ui.prompt('請輸入表單預計存放的雲端位置連結').getResponseText();
  let destinationFolderId = destinationFolderUrl.split('/')[5];
  if (destinationFolderUrl.split('/')[4] != "folders"){
    Logger.log("請填寫雲端資料夾位置之連結")
  }
  moveFile(new_form_2_id, destinationFolderId);
  
  SetForm2(new_form_2, yr)
}

function moveFile(fileId, destinationFolderId) {
  let destinationFolder = DriveApp.getFolderById(destinationFolderId);
  DriveApp.getFileById(fileId).moveTo(destinationFolder);
}

function SetForm2(form, yr){
  form.setCollectEmail(true);

  let sht = SpreadsheetApp.getActiveSheet();
  let startRow = 4;
  let numRows = sht.getLastRow() - startRow +1;
  let startCol = 2;
  let numCols = sht.getLastColumn() - startCol +1;
  let rg = sht.getRange(startRow, startCol, numRows, numCols);
  let dt = rg.getValues();

  let crs = [];
  for (let i in dt) {
    if(dt[i][8]==(yr+'-2')){ // 若是下學期課程才會執行
      crs.push(dt[i][0].concat(' ', dt[i][6]))
    }
  }
  let list_item = form.addListItem();
  list_item.setTitle('請選擇您的課程編號')
          .setChoiceValues(crs)
          .setRequired(true);

  // 第一點
  form.addPageBreakItem().setTitle('1. 經費執行情形(含自籌款)');
  form.addSectionHeaderItem().setTitle('原核定計畫金額');
  form.addTextItem().setTitle('人事費').setRequired(true);
  form.addTextItem().setTitle('業務費').setRequired(true);
  form.addTextItem().setTitle('設備費').setRequired(true);

  form.addSectionHeaderItem().setTitle('目前實支數');
  form.addTextItem().setTitle('人事費').setRequired(true);
  form.addTextItem().setTitle('業務費').setRequired(true);
  form.addTextItem().setTitle('設備費').setRequired(true);

  // 第二點
  form.addPageBreakItem().setTitle('2. 教學設備採購進度')
                        .setHelpText('格式:品項*個數/金額 EX: 伺服器*1/NT$20,000');
  form.addParagraphTextItem().setTitle('預計購買項目(含金額)').setRequired(true);
  form.addParagraphTextItem().setTitle('已完成招標/完成請購之項目(含金額)').setRequired(true);

  // 第三點
  form.addPageBreakItem().setTitle('3. 課程與模組結合使用情形')
  // form.addCheckboxItem().setTitle('課程之採用模組')
  //                       .setChoiceValues(md)
  //                       .setRequired(true);
  // form.addTextItem().setTitle('採用模組時數\n例如: A-1 (12小時)、C-2 (9小時)').setRequired(true);
  form.addTextItem().setTitle('請列出課程大綱並註明重點模組採用總時數，並將檔案連結貼於此').setRequired(true);
  //var img = UrlFetchApp.fetch('https://drive.google.com/file/d/1Ea3Y3aQ8yF8k8f3oPFq0345HLlAegr39/view?usp=drive_link')
  //form.addImageItem().setTitle('範例').setImage(img);

  // 第四點
  form.addPageBreakItem().setTitle('4. 課程開授成效');
  form.addSectionHeaderItem().setTitle('申請書內目標值');
  form.addTextItem().setTitle('修課人次\nEX: 大學部(5人)、碩士班(20人)、博士班(0人)')
                    .setRequired(true);
  form.addTextItem().setTitle('專題作品數').setRequired(true);
  form.addParagraphTextItem().setTitle('質化成效說明').setRequired(true);

  form.addSectionHeaderItem().setTitle('目前已達成值');
  form.addTextItem().setTitle('修課人次\nEX: 大學部(5人)、碩士班(20人)、博士班(0人)')
                    .setRequired(true);
  form.addTextItem().setTitle('專題作品數').setRequired(true);
  form.addParagraphTextItem().setTitle('質化成效說明').setRequired(true);
  form.addParagraphTextItem().setTitle('未達目標值，須在此說明理由')
                             .setHelpText("例如: 跟必修課程時間重疊...");

  // 第五點
  form.addPageBreakItem().setTitle('5. 參與聯盟活動、競賽情形');
  form.addSectionHeaderItem().setTitle('申請書內目標值');
  form.addTextItem().setTitle('參與聯盟相關課程推廣研習、座談之人次').setRequired(true);
  form.addTextItem().setTitle('參與聯盟相關課程推廣研習、座談之場次').setRequired(true);
  form.addTextItem().setTitle('參與聯盟相關競賽學生人數').setRequired(true);
  form.addParagraphTextItem().setTitle('其他');

  form.addSectionHeaderItem().setTitle('目前已達成值');
  form.addTextItem().setTitle('參與聯盟相關課程推廣研習、座談之人次').setRequired(true);
  form.addTextItem().setTitle('參與聯盟相關課程推廣研習、座談之場次').setRequired(true);
  form.addTextItem().setTitle('參與聯盟相關競賽學生人數').setRequired(true);
  form.addParagraphTextItem().setTitle('其他');

  // 第六點
  form.addPageBreakItem().setTitle('6. 業界或校外講師參與教學情形');
  form.addParagraphTextItem().setTitle('業界專家學者及其他跨領域教師實際參與計畫情形');
}

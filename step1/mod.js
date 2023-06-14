function createNewForm(){
    let new_form = FormApp.create("模組期中書面審查意見表test");
    let new_form_id = new_form.getId();
    let destinationFolderId = "1ZD7uUQQdnLBXjLe0D7uoEtULw8fXBBdP"; //須根據情況修改
    moveFile(new_form_id, destinationFolderId);
    
    SetForm(new_form)
  }
  
  function moveFile(fileId, destinationFolderId) {
    let destinationFolder = DriveApp.getFolderById(destinationFolderId);
    DriveApp.getFileById(fileId).moveTo(destinationFolder);
  }
  
  function SetForm(form){
    var cbValidation = FormApp.createCheckboxValidation()
    .requireSelectAtLeast(1)
    .build();
  
    let des = '111年度模組期中報告書：\n\
  https://drive.google.com/drive/folders/14Z0kuttyDbrK6RiuNXYtRditDFw5BUfG\n\
  \n\
  說明如下：\n\
  \n\
  一、 考評重點：\n\
  \n\
      1.整體模組教材開發、試教、推廣情形\n\
  \n\
      2.公開徵件模組教材規劃是否合宜\n\
  \n\
      3.模組教材績效達成情形\n\
  \n\
      4.模組教材經費使用情況\n\
  \n\
  \n\
  二、 評等分數： 10:極優, 9:優, 8:良, 7:尚可, 6:可, 5:普通, 4:略差, 3:差, 2:極差, 1:劣';
  
    form.setDescription(des);
    form.setCollectEmail(true);
  
    let ls = ['呂良鴻 委員', '吳安宇 委員', '許明華 委員', '鄭國興 委員', '張振豪 委員'];
    let list_item = form.addListItem();
    list_item.setTitle('請選擇您的身分')
            .setChoiceValues(ls)
            .setRequired(true);
    
    let sht = SpreadsheetApp.getActiveSheet();
    let startRow = 4;
    let numRows = sht.getLastRow() - startRow +1;
    let startCol = 1;
    let numCols = sht.getLastColumn() - startCol +1;
    let rg = sht.getRange(startRow, startCol, numRows, numCols);
    let dt = rg.getValues();
  
    let cb_ls = ['教材開發不如預期', '業務費執行率偏低', '設備費執行率偏低', '報告內容說明不夠詳盡', '無(於下一題說明)'];
  
    for (let i in dt) {
      //dt[i] is a row
      add_pages(form, dt[i])
      for (let j = 1; j <= 4; j++){ // 增加審查重點， 因為有四項，所以j<=4
        add_list(form, dt[i], j)
      }
  
      let cb = form.addCheckboxItem();
      cb.setTitle(dt[i][0]+' 綜合審查意見(可複選或在下欄中填寫補充意見)')
          .setChoiceValues(cb_ls)
          .setValidation(cbValidation)
          .setRequired(true);
  
      let dt_qus = form.addParagraphTextItem();
      dt_qus.setTitle(dt[i][0]+' 審查意見補充說明')
            .setRequired(true);
  
      let mc_ls = [10,9,8,7,6,5,4,3,2,1]
      let mc = form.addMultipleChoiceItem();
      mc.setTitle(dt[i][0]+' 綜合評分')
        .setHelpText('10:極優, 9:優, 8:良, 7:尚可, 6:可, 5:普通, 4:略差, 3:差, 2:極差, 1:劣')
        .setChoiceValues(mc_ls)
        //.showOtherOption(true)
        .setRequired(true);
    }
    
  }
  
  function add_pages(form, row){
    let pg = form.addPageBreakItem().setTitle('智慧終端裝置晶片系統與應用聯盟');
    pg.setHelpText('課程名稱: '+ row[4] 
    +'\n計畫編號: '+ row[0] 
    +'\n模組教師: '+ row[3] 
    +'\n學校/系所(服務單位): '+row[1]+'/'+row[2])
  }
  
  function add_list(form, row, sel){
    let lst = form.addListItem();
    if(sel==1){
      lst.setTitle(row[0]+' 審查重點- 第一項：整體模組教材開發、試教、推廣情形')
    }
    else if(sel==2){
      lst.setTitle(row[0]+' 審查重點- 第二項：公開徵件模組教材規劃是否合宜')
    }
    else if(sel==3){
      lst.setTitle(row[0]+' 審查重點- 第三項：模組教材績效達成情形')
    }
    else{
      lst.setTitle(row[0]+' 審查重點- 第四項：模組教材經費使用情況')
    }
  
    let ls = ['優', '佳', '尚可', '不佳'];
    lst.setChoiceValues(ls)
        .setRequired(true);
  }
  
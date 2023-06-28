function createFolderLink() {
    var cur_fd_id = get_current_folder_id();
  
    let sht = SpreadsheetApp.getActiveSheet();
    let numRows = sht.getLastRow();
    let numCols = sht.getLastColumn();
    let rg = sht.getRange(1, 1, numRows, numCols);
    let dt = rg.getValues();
  
    // 存放報告書的資料夾，每次會變動，要修改
    let rp_fd = DriveApp.getFolderById('17XUdqzV0VP2laf1whi4_zq7i4_depmgu')
    var arr = []
    for (let i in dt) {
      //dt[i] is a row
  
      // 創建雲端資料夾，以委員姓名命名
      if (i==0) {
        for (let d in dt[0]) {
          let nm = dt[0][d]
          if (nm!='') {
            var folder = DriveApp.createFolder(nm+'委員');
            arr.push(folder.getId());
  
            //將創建出的資料夾 移動到當前雲端資料夾
            folder.moveTo(DriveApp.getFolderById(cur_fd_id));
          }
        }
      }
      else {
        var files = rp_fd.searchFiles("title contains '" + dt[i][0] + "'");
        var file = files.next();
        for (var stp=1; stp<numCols; stp++) {
          if (dt[i][stp]=='v') {
            file.makeCopy(DriveApp.getFolderById(arr[stp-1]));
          }
        }
      }
    }
    console.log(arr);
}
  
function get_current_folder_id() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var current = DriveApp.getFileById(ss.getId());
    var fdrs = current.getParents();

    return fdrs.next().getId();
}
  
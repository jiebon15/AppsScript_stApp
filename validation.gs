function applyValidationToCell(list,cell){
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(list).build();
  cell.setDataValidation(rule);
}

function validateAlamat(upi){
  lastRowData = wsData2.getLastRow();
  var isDataExists = false;
  for (var i = 1; i <= lastRowData; i++){
    if (upi == wsData2.getRange(i, 1).getValue()){
      isDataExists = true;
      break;
    }
  }
  if (!isDataExists)
  {
    var blankRow=wsData2.getLastRow()+1;
    wsData2.getRange(blankRow, 1).setValue(upi);
    wsData2.getRange(blankRow, 2).setValue(alamatupi);
  }
}

function validateEntry(){
  lastRowData = datasheet.getLastRow();
  var noStExists = false;
  for (var i = 1; i <= lastRowData; i++){
    if (nost == datasheet.getRange(i, 1).getValue()){
      noStExists = true;
    }
  }
  if(noStExists == true){
    SpreadsheetApp.getUi().alert("Nomor Surat Sudah Dipakai");
    return false;
  }
  else if(ws.getRange("C5").isBlank()==true){
    SpreadsheetApp.getUi().alert("Silahkan Masukan Jenis Surat");
    return false;
  }
  else if(ws.getRange("C7").isBlank()==true){
    SpreadsheetApp.getUi().alert("Silahkan Masukan Nomor Surat");
    return false;
  }
  else if(ws.getRange("C9").isBlank()==true){
    SpreadsheetApp.getUi().alert("Silahkan Masukan Tanggal Surat");
    return false;
  }
  else if(ws.getRange("F7").isBlank()==true){
    SpreadsheetApp.getUi().alert("Silahkan Masukan Yang Menandatangani Surat");
    return false;
  }
  else if(ws.getRange("I7").isBlank()==true){
    SpreadsheetApp.getUi().alert("Silahkan Masukan Jabatan Yang Menandatangani Surat");
    return false;
  }
  else if(ws.getRange("C12").isBlank()==true){
    SpreadsheetApp.getUi().alert("Silahkan Masukan Nama Ketua");
    return false;
  }
  else if(ws.getRange("C16").isBlank()==true){
    SpreadsheetApp.getUi().alert("Silahkan Masukan Tanggal Kegiatan");
    return false;
  }
  else if(ws.getRange("C18").isBlank()==true){
    SpreadsheetApp.getUi().alert("Silahkan Masukan UPI / UUP");
    return false;
  }
  else if(ws.getRange("C20").isBlank()==true){
    SpreadsheetApp.getUi().alert("Silahkan Masukan Alamat UPI / UUP");
    return false;
  }
  return true;
}

/////////////// General Variabel \\\\\\\\\\\\\\\\\\\\\\\\\
var mainWsName = "main";
var ws = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(mainWsName);
var wsData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("data");
var wsData2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("UPI");
var datasheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("database");

const docTempInspeksiCKIB = DriveApp.getFileById("//replace\\"); //ganti dengan id template doc
const docTempInspeksiCKIBTahunan = DriveApp.getFileById("//replace\\");
const docTempMonsurCKIB = DriveApp.getFileById("//replace\\");
const docTempInspeksiHACCP = DriveApp.getFileById("//replace\\");
const docTempMonsurImpor = DriveApp.getFileById("//replace\\");
const docTempOfficialControl = DriveApp.getFileById("//replace\\");
const docTempPenilaianIKI = DriveApp.getFileById("//replace\\");
const docTempPerpanjanganIKI = DriveApp.getFileById("//replace\\");
const docTempVerifikasiHACCP = DriveApp.getFileById("//replace\\");
const docTempMonevHACCP = DriveApp.getFileById("//replace\\");
const docTempTraceability = DriveApp.getFileById("//replace\\");
const docTempSurveillanceHACCP = DriveApp.getFileById("//replace\\");
const docTempCPIB = DriveApp.getFileById("//replace\\");
const docTempRegUPI = DriveApp.getFileById("//replace\\");

const docFile = DriveApp.getFolderById("//replace\\"); //ganti dengan id target folder

var st = ws.getRange("C5").getValue();
var nost = ws.getRange("C7").getValue();
var ketua = ws.getRange("C12").getValue();
var anggota1 = ws.getRange("C14").getValue();
var anggota2 = ws.getRange("F14").getValue();
var upi = ws.getRange("C18").getValue();
var alamatupi = ws.getRange("C20").getValue();
var tglkegiatan = ws.getRange("C16").getValue();
var blnkegiatan = ws.getRange("F16").getValue();
var thnkegiatan = ws.getRange("I16").getValue();
var tglsurat = ws.getRange("C9").getValue();
var blnsurat = ws.getRange("F9").getValue();
var thnsurat = ws.getRange("I9").getValue();
var jabatan = ws.getRange("I7").getValue();
var ttd = ws.getRange("F7").getValue();
var rom = romawi(blnsurat);
//////////////////////////////////////\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\


///////////////////// Main Function Generate Button \\\\\\\\\\\\\\\\\\\\\
function generate(){
  if (st === "Inspeksi CKIB Tahunan"){
    var noreg = true;
    var noregketua = noRegKarantina(ketua);
    var noreganggota1 = noRegKarantina(anggota1);
    //var noreganggota2 = noRegKarantina(anggota2);
    merge(docTempInspeksiCKIBTahunan, noreg, noregketua, noreganggota1);
  }
  else if (st === "Inspeksi CKIB"){
    var noreg = true;
    var noregketua = noRegKarantina(ketua);
    var noreganggota1 = noRegKarantina(anggota1);
    //var noreganggota2 = noRegKarantina(anggota2);
    merge(docTempInspeksiCKIB, noreg, noregketua, noreganggota1);
  }
  else if (st === "Monsur CKIB"){
    var noreg = true;
    var noregketua = noRegKarantina(ketua);
    var noreganggota1 = noRegKarantina(anggota1);
    //var noreganggota2 = noRegKarantina(anggota2);
    merge(docTempMonsurCKIB, noreg, noregketua, noreganggota1);
  }
  else if (st === "Penilaian IKI"){
    var noreg = true;
    var noregketua = noRegKarantina(ketua);
    var noreganggota1 = noRegKarantina(anggota1);
    //var noreganggota2 = noRegKarantina(anggota2);
    merge(docTempPenilaianIKI, noreg, noregketua, noreganggota1);
    SpreadsheetApp.getUi().alert("Silahkan Buka CKIB Online  Untuk Melanjutkan")
  }
  else if (st === "Perpanjangan IKI"){
    var noreg = true;
    var noregketua = noRegKarantina(ketua);
    var noreganggota1 = noRegKarantina(anggota1);
    //var noreganggota2 = noRegKarantina(anggota2);
    merge(docTempPerpanjanganIKI, noreg, noregketua, noreganggota1);
    SpreadsheetApp.getUi().alert("Silahkan Buka CKIB Online Untuk Melanjutkan")
  }
  else if (st === "Surveillance HACCP"){
    var noreg = true;
    var noregketua = noRegMutu(ketua);
    var noreganggota1 = noRegMutu(anggota1);
    merge(docTempSurveillanceHACCP, noreg, noregketua, noreganggota1);
    SpreadsheetApp.getUi().alert("Silahkan Buka Honest Untuk Melanjutkan");
  }
  else if (st === "Official Control"){
    var noreg = true;
    var noregketua = noRegMutu(ketua);
    merge(docTempOfficialControl, noreg, noregketua);
  }
  else if (st === "Inspeksi HACCP"){
    var noreg = true;
    var noregketua = noRegMutu(ketua);
    var noreganggota1 = noRegMutu(anggota1);
    merge(docTempInspeksiHACCP, noreg, noregketua, noreganggota1);
  }
  else if (st === "Verifikasi HACCP"){
    var noreg = true;
    var noregketua = noRegMutu(ketua);
    var noreganggota1 = noRegMutu(anggota1);
    merge(docTempVerifikasiHACCP, noreg, noregketua, noreganggota1);
  }
  else if(st === "Monsur Impor"){
    var noreg = false;
    merge(docTempMonsurImpor, noreg);
  }
  else if(st === "Monev HACCP"){
    var noreg = true;
    var noregketua = noRegMutu(ketua);
    var noreganggota1 = noRegMutu(anggota1);
    merge(docTempMonevHACCP, noreg, noregketua, noreganggota1);
  }
  else if(st === "Traceability"){
    var noreg = true;
    var noregketua = noRegMutu(ketua);
    var noreganggota1 = noRegMutu(anggota1);
    merge(docTempTraceability, noreg, noregketua, noreganggota1);
  }
  else if(st === "CPIB"){
    var noreg = false;
    merge(docTempCPIB, noreg);
  }
  else if(st === "Reg UPI"){
    var noreg = true;
    var noregketua = noRegMutu(ketua);
    var noreganggota1 = noRegMutu(anggota1);
    merge(docTempRegUPI, noreg, noregketua, noreganggota1);
  }
}
//////////////////////////////////////\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

//////////////////////////////////////////Function Find and Replace\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
function merge(temp, noreg, noregketua, noreganggota1) {
  //if (validateEntry()==true) {
       
    //nipanggota2 = nip(anggota2);
    var nama = nost + " - " + st +" "+ upi;
    const tempFile = temp.makeCopy(docFile).setName(nama);
    const tempDocFile = DocumentApp.openById(tempFile.getId());
    const body = tempDocFile.getBody();
    body.replaceText("{rom}", rom);
    body.replaceText("{nost}", nost);
    body.replaceText("{upi}", upi);
    body.replaceText("{alamatupi}", alamatupi);
    body.replaceText("{tglkegiatan}", tglkegiatan);
    body.replaceText("{blnkegiatan}", blnkegiatan);
    body.replaceText("{thnkegiatan}", thnkegiatan);
    body.replaceText("{tglsurat}", tglsurat);
    body.replaceText("{blnsurat}", blnsurat);
    body.replaceText("{thnsurat}", thnsurat);
    body.replaceText("{jabatan}", jabatan);
    body.replaceText("{ttd}", ttd);
    if (noreg == true){
      body.replaceText("{ketua}", ketua);
      nipketua = nip(ketua);
      body.replaceText("{nipketua}", nipketua);
      /*if (ws.getRange("F14").isBlank()==false){
        body.replaceText("{anggota1}", anggota1);
        body.replaceText("{nipanggota1}", nipanggota1);
        body.replaceText("{anggota2}", anggota2);
        body.replaceText("{nipanggota2}", nipanggota2);
        body.replaceText("{noregketua}", noregketua);
        body.replaceText("{noreganggota1}", noreganggota1);
        body.replaceText("{noreganggota2}", noreganggota2);
      } */
      if (ws.getRange("C14").isBlank()==false){
        body.replaceText("{anggota1}", anggota1);
        nipanggota1 = nip(anggota1);
        body.replaceText("{nipanggota1}", nipanggota1);
        body.replaceText("{noregketua}", noregketua);
        body.replaceText("{noreganggota1}", noreganggota1);
      }
      else{
        body.replaceText("{noregketua}", noregketua);
      }
    }
  else{
      body.replaceText("{ketua}", ketua);
      nipketua = nip(ketua);
      body.replaceText("{nipketua}", nipketua);
      if (ws.getRange("C14").isBlank()==false){
        body.replaceText("{anggota1}", anggota1);
        nipanggota1 = nip(anggota1);
        body.replaceText("{nipanggota1}", nipanggota1)
      }
      /*else if (ws.getRange("F14").isBlank()==false){
        body.replaceText("{anggota1}", anggota1);
        body.replaceText("{nipanggota1}", nipanggota1);
        body.replaceText("{anggota2}", anggota2);
        body.replaceText("{nipanggota2}", nipanggota2);
      }*/
    }
    tempDocFile.saveAndClose();
    clear();
  //}
}
//////////////////////////////////////\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\


//////////////////////////////////////////Function Input Botton\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
function submitData() {
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert("Submit", 'Apakah Data Sudah Benar?',ui.ButtonSet.YES_NO);

  // Checking the user response and proceed with clearing the form if user selects Yes
  if (response == ui.Button.NO) 
  {return;//exit from this function
  } 

  if (validateEntry()==true) {
    validateAlamat(upi);
    var blankRow=datasheet.getLastRow()+1; //identify the next blank row
    var currentUser = Session.getEffectiveUser();
    var userFullName = currentUser.getEmail();
    datasheet.getRange(blankRow, 1).setValue(nost);
    datasheet.getRange(blankRow, 2).setValue(st);
    datasheet.getRange(blankRow, 3).setValue(rom);
    datasheet.getRange(blankRow, 4).setValue(tglkegiatan);
    datasheet.getRange(blankRow, 5).setValue(blnkegiatan);
    datasheet.getRange(blankRow, 6).setValue(thnkegiatan);
    datasheet.getRange(blankRow, 7).setValue(upi);
    datasheet.getRange(blankRow, 8).setValue(ketua);
    datasheet.getRange(blankRow, 9).setValue(anggota1);
    datasheet.getRange(blankRow, 10).setValue(anggota2); 
    datasheet.getRange(blankRow, 11).insertCheckboxes();
    datasheet.getRange(blankRow, 12).insertCheckboxes();
    datasheet.getRange(blankRow, 13).insertCheckboxes();
    datasheet.getRange(blankRow, 15).setValue(userFullName);
    var response2 = ui.alert("Generate", 'Buat Document?',ui.ButtonSet.YES_NO);
    if (response2 == ui.Button.NO) {
      return;
    }
    else {
      generate();
      clear();
    }
 }
}
//////////////////////////////////////\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

//////////////////////////////////////////Function Clear Botton\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
function clear() {
  ws.getRange("C5").clearContent();
  ws.getRange("C12").clearContent();
  ws.getRange("C14").clearContent();
  ws.getRange("F14").clearContent();
  ws.getRange("C16").clearContent();
  ws.getRange("C18").clearContent();
  ws.getRange("C20").clearContent();
  ws.getRange("C20").clearDataValidations();
}
//////////////////////////////////////\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

//////////////////////////////////////////Function Search Botton\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
function cari(){
  applyUPIValidation();
}
//////////////////////////////////////\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

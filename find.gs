function applyUPIValidation(){
    var returnValues = [];
    val = ws.getRange("C18").getValue();
    ws.getRange("C20").clearContent(); 
    lastRowData = wsData2.getLastRow();
    for(var i = 1; i <= lastRowData; i++)
    {
     if(val == wsData2.getRange(i, 1).getValue())
     {
        returnValues.push(wsData2.getRange(i, 2).getValue());      
     }
    }
    var cell = ws.getRange("C20");
    applyValidationToCell(returnValues,cell);
}

function nip(nama){
  var returnValue;
  lastRowData = wsData.getLastRow();
  for(var i = 2; i <= lastRowData; i++)
  {
    if(nama == wsData.getRange(i, 1).getValue())
    {
      returnValue = wsData.getRange(i, 2).getValue();      
    }
  }
  return returnValue;
}

function noRegKarantina(nama){
  var returnValue;
  lastRowData = wsData.getLastRow();
  for(var i = 2; i <= lastRowData; i++)
  {
    if(nama == wsData.getRange(i, 1).getValue())
    {
      returnValue = wsData.getRange(i, 4).getValue();      
    }
  }
  return returnValue;
}

function noRegMutu(nama){
  var returnValue;
  lastRowData = wsData.getLastRow();
  for(var i = 2; i <= lastRowData; i++)
  {
    if(nama == wsData.getRange(i, 1).getValue())
    {
      returnValue = wsData.getRange(i, 5).getValue();      
    }
  }
  return returnValue;
}

function romawi(indonesianMonth) {
  var romawiMapping = {
    "Januari": "I",
    "Februari": "II",
    "Maret": "III",
    "April": "IV",
    "Mei": "V",
    "Juni": "VI",
    "Juli": "VII",
    "Agustus": "VIII",
    "September": "IX",
    "Oktober": "X",
    "November": "XI",
    "Desember": "XII"
  };
  return romawiMapping[indonesianMonth] || indonesianMonth;
}
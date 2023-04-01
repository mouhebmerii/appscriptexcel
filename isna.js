function countall() {
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheets = spreadsheet.getSheets();
  var enAttenteCount = [];
  var enCoursCount = [];
  var okCount = [];
  var blocCount = [];
  var koCount = [];
  var naCount = [];
  
  var sheetNames = [];


  
  for (var i=0;i<sheets.length;i++){
  
      var sheetName = sheets[i].getName();
      if (sheetName.indexOf('<') === -1 && sheetName.indexOf('>') === -1){
      sheetNames.push([sheetName]);
    }
  }
  var targetRange = spreadsheet.getRange("A3:A" + (sheetNames.length+2));
  
  targetRange.setValues(sheetNames);



var values = [];
for (var i=0;i<sheets.length;i++){
    var sheetName = sheets[i].getName();
    if (sheetName.indexOf('<') === -1 && sheetName.indexOf('>') === -1){
var value = sheets[i].getRange("B2").getValue();
values.push([value]);

    }
}
var targetRange = spreadsheet.getRange("B3:B" + (values.length+2));

targetRange.setValues(values);


var valuest = [];
for (var i=0;i<sheets.length;i++){
   var sheetName = sheets[i].getName();
   if (sheetName.indexOf('<') === -1 && sheetName.indexOf('>') === -1){
var value = sheets[i].getRange("C2").getValue();
valuest.push([value]);

   }
}
var targetRange = spreadsheet.getRange("C3:C" + (valuest.length+2));

targetRange.setValues(valuest);





var valuesS = [];
for (var i=0;i<sheets.length;i++){
    var sheetName = sheets[i].getName();
    if (sheetName.indexOf('<') === -1 && sheetName.indexOf('>') === -1){
var value = sheets[i].getRange("F2").getValue();
valuesS.push([value]);

    }
}
var targetRange = spreadsheet.getRange("D3:D" + (valuesS.length+2));

targetRange.setValues(valuesS);



var count = [];
  for (var i=0;i<sheets.length;i++){
    var sheetName = sheets[i].getName();
    if (sheetName.indexOf('<') === -1 && sheetName.indexOf('>') === -1){
      var data = sheets[i].getDataRange().getValues();
      var ctCount = 0;
      for (var j=2; j<data.length; j++){
        if (data[j][0].toString().trim() !== ""){
          ctCount++;
        }
      }
      count.push([ctCount]);
    }
  }
  var targetRange = spreadsheet.getRange("E3:E" + (count.length+2));
  targetRange.setValues(count);








var echecc = [];
var ctExec = [];
var avcCount = [];
var rafCount = [];
var ceCount = 0;
  for (var i = 0; i < sheets.length; i++) {
    var sheetName = sheets[i].getName();
    if (sheetName.indexOf('<') === -1 && sheetName.indexOf('>') === -1) {


      var data = sheets[i].getDataRange().getValues();
      var ceCount = 0;
      for (var j=2; j<data.length; j++){
        if (data[j][0].toString().trim() !== ""){
          ceCount++;
          Logger.log(ceCount)
        }
      }










      var data = sheets[i].getDataRange().getValues();
      var enAttenteCt = 0;
      var enCoursCt = 0;
      var okCt = 0;
      var bCt = 0;
      var koCT = 0;
      var nact = 0;
      var exc = 0;
      var ec = 0;
      var avc = 0;
      var raf = 0;

      for (var j = 0; j < data.length; j++) {
        if (data[j][5] === "En attente") {
          enAttenteCt++;
        } else if (data[j][5] === "En cours") {
          enCoursCt++;
        } else if (data[j][5] === "OK") {
          okCt++;
        } else if (data[j][5] === "Bloqué") {
          bCt++;
        } else if (data[j][5] === "KO") {
          koCT++;
        } else if (data[j][5] === "NA") {
          nact++;
        }

        exc = okCt + koCT;
       
    // var count = [];

    
        // var ctc = count[i];
    if (parseInt(ceCount)!==0){
        ec = (koCT / parseInt(ceCount)) * 100;

        var percentage =(ec + "%");
        avc = (okCt / parseInt(ceCount)) * 100;
        var avcper =(avc + "%");
        

        raf = (enCoursCt+bCt+koCT / parseInt(ceCount)) *100;
        var rafperccentage = (raf +"%");
    } else{
percentage = "0%";
avcper = "0%";
rafperccentage = "0%";
    }



      }
      enAttenteCount.push([enAttenteCt]);
      enCoursCount.push([enCoursCt]);
      okCount.push([okCt]);
      blocCount.push([bCt]);
      koCount.push([koCT]);
      naCount.push([nact]);
      echecc.push([percentage]);
      avcCount.push([avcper]);
      rafCount.push([rafperccentage]);
  

      ctExec.push([exc]);

    }
  }

  var enAttenteTargetRange = spreadsheet.getRange("F3:F" + (enAttenteCount.length + 2));
  enAttenteTargetRange.setValues(enAttenteCount);

  var enCoursTargetRange = spreadsheet.getRange("G3:G" + (enCoursCount.length + 2));
  enCoursTargetRange.setValues(enCoursCount);

  var okTargetRange = spreadsheet.getRange("H3:H" + (okCount.length + 2));
  okTargetRange.setValues(okCount);

  var okTargetRange = spreadsheet.getRange("I3:I" + (blocCount.length + 2));
  okTargetRange.setValues(blocCount);

  var okTargetRange = spreadsheet.getRange("J3:J" + (koCount.length + 2));
  okTargetRange.setValues(koCount);

  var okTargetRange = spreadsheet.getRange("K3:K" + (naCount.length + 2));
  okTargetRange.setValues(naCount);

  var okTargetRange = spreadsheet.getRange("M3:M" + (ctExec.length + 2));
  okTargetRange.setValues(ctExec);

  
  var okTargetRange = spreadsheet.getRange("N3:N" + (echecc.length + 2));
  okTargetRange.setValues(echecc);
    
  var okTargetRange = spreadsheet.getRange("O3:O" + (avcCount.length + 2));
  okTargetRange.setValues(avcCount);
      
  var okTargetRange = spreadsheet.getRange("P3:P" + (rafCount.length + 2));
  okTargetRange.setValues(rafCount);
}

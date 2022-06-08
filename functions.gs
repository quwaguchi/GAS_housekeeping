function rewriteSpreadsheet(){

  //現在月のシートがあるか確認。無い場合は左から４番目に設置し、概要シートに反映する。
  var currentMonth = Utilities.formatDate(new Date(),'JST','yyyy/MM')
  if (as.getSheets()[3].getSheetName() != currentMonth){
    
    tempSheet.copyTo(as).setName(currentMonth).activate()
    as.moveActiveSheet(4)

    var refTemp = "='"+currentMonth+"'!"
    totalSheet.appendRow([currentMonth,refTemp+"D2",refTemp+"E2",refTemp+"F2",refTemp+"G2"])
  }

  //支払月と同じシートを探す
  var sheetsNames = as.getSheets().map(sheet => sheet.getName())
  var paidMonth = Utilities.formatDate(paidDate, 'JST', 'yyyy/MM')
  var paidMonthIndex = sheetsNames.indexOf(paidMonth)
  var paidMonthSheet = as.getSheetByName(sheetsNames[paidMonthIndex])

  //結,晴人の負担額を計算
  if (chargeYui==="" && chargeHaru===""){
    chargeYui = paidAmount/2
    chargeHaru = paidAmount/2
  }else if(chargeYui===""){
    chargeYui = paidAmount - chargeHaru
  }else if(chargeHaru===""){
    chargeHaru = paidAmount - chargeYui
  }

  //該当シートに最新回答を反映
  if (chargeYui+chargeHaru==paidAmount && payer=="結"){
    paidMonthSheet.appendRow([payer,paidDate,purpose,paidAmount,,chargeYui,chargeHaru])
  }else if(chargeYui+chargeHaru==paidAmount && payer=="晴人"){
    paidMonthSheet.appendRow([payer,paidDate,purpose,,paidAmount,chargeYui,chargeHaru])
  }else{ //負担額に矛盾がある場合
    respSheet.getRange(respLR,respLC+1).setValue("エラー！")
  }

}
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

  //w,mの負担額を計算
  if (charge_w==="" && charge_m===""){
    charge_w = paidAmount/2
    charge_m = paidAmount/2
  }else if(charge_w===""){
    charge_w = paidAmount - charge_m
  }else if(charge_m===""){
    charge_m = paidAmount - charge_w
  }

  //該当シートに最新回答を反映
  if (charge_w+charge_m==paidAmount && payer=="w"){
    paidMonthSheet.appendRow([payer,paidDate,purpose,paidAmount,,charge_w,charge_m])
  }else if(charge_w+charge_m==paidAmount && payer=="m"){
    paidMonthSheet.appendRow([payer,paidDate,purpose,,paidAmount,charge_w,charge_m])
  }else{ //負担額に矛盾がある場合
    respSheet.getRange(respLR,respLC+1).setValue("エラー！")
  }

}

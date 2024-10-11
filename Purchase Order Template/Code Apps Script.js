function AddItem()
{
  
  //DEFINE ALL ACTIVE SHEETS
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  //DEFINE MENU SHEET          
  var poSheet = ss.getSheetByName("POS");
  var itemSheet = ss.getSheetByName("ITEMS");
  
  //GET NEXT ROW OF PO SHEET
  var lastrowPO = poSheet.getLastRow() + 1;
  
  //GET LAST ROW OF ITEM SHEET
  var lastrowItem = itemSheet.getLastRow();
  
  // GET VALUE OF PART AND QUANTITY
  var part = poSheet.getRange('B13').getValue();
  var quantity = poSheet.getRange('B14').getValue();
  
  // GET UNIT PRICE FROM ITEM SHEET
  for(var i = 2; i <= lastrowItem; i++)
  {
    if(part == itemSheet.getRange(i, 1).getValue())
    {
      var description = itemSheet.getRange(i, 2).getValue();
      var unitCost = itemSheet.getRange(i, 3).getValue();
    }
  }
  
  // POPULATE PO SHEET
  poSheet.getRange(lastrowPO, 1).setValue(part);
  poSheet.getRange(lastrowPO, 2).setValue(description);
  poSheet.getRange(lastrowPO, 3).setValue(quantity);
  poSheet.getRange(lastrowPO, 4).setValue(unitCost).setNumberFormat("$#,###.00");
  
}


function createPO()
{
  //DEFINE ALL ACTIVE SHEETS
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  //DEFINE MENU SHEET          
  var poSheet = ss.getSheetByName("POS");
  var vendorSheet = ss.getSheetByName("VENDORS");
  var settingSheet = ss.getSheetByName("SETTINGS");
  var printSheet = ss.getSheetByName("PRINT PO");
  
  //GET VALUES
  var name = poSheet.getRange(6,2).getValue();
  var invoice_number = poSheet.getRange(7,2).getValue();
  var ship_date = poSheet.getRange(8,2).getValue();
  var ship_via = poSheet.getRange(9,2).getValue();
  var terms = poSheet.getRange(10,2).getValue();
  var ship_and_handle = poSheet.getRange(11,2).getValue();
  var po_number = settingSheet.getRange(1,2).getValue();
  var next_po_number = po_number + 1;
  settingSheet.getRange(1,2).setValue(next_po_number);
  
  // GET VENDOR LAST ROW
  var lastrowVendor = vendorSheet.getLastRow();
  
  // GET VENDOR FIELDS
  for(var i = 2; i <= lastrowVendor; i++)
  {
    if(name == vendorSheet.getRange(i, 1).getValue())
    {
      var companyName = vendorSheet.getRange(i,2).getValue();
      var streetAddress = vendorSheet.getRange(i,3).getValue();
      var city = vendorSheet.getRange(i,4).getValue();
      var state = vendorSheet.getRange(i,5).getValue();
      var zip = vendorSheet.getRange(i,6).getValue();
      var phone_number = vendorSheet.getRange(i,7).getValue();
      var email = vendorSheet.getRange(i,8).getValue();
      var tax_rate = vendorSheet.getRange(i,9).getValue();
    }
  }
  
  // SET PO DATE
  var currentDate = new Date();
  var currentMonth = currentDate.getMonth()+1;
  var currentYear = currentDate.getFullYear();
  var date = currentMonth.toString() + '/' + currentDate.getDate().toString() + '/' + currentYear.toString();

  // GET LAST ROW OF PRINT SHEET
  var lastrowPrint = printSheet.getLastRow();
  
  // FIND HOW MANY ITEMS ROWS TO DELETE
  var x_count = 0
  for(var v = 27; v <= lastrowPrint; v++)
  {
    
    if(printSheet.getRange(v, 6).getValue() != 'Subtotal')
    {  
      x_count++;
    }
    else
    {
      break;
    }
  }
  
  var lastrowPrint = 27 + x_count;
  
  //Logger.log(lastrowPrint);
  
  // DELETE ITEMS ROWS FROM PO
  if((lastrowPrint - 27) != 0)
  {
    printSheet.deleteRows(27, lastrowPrint - 27);
  }  
  
  // SET VALUES ON PO  
  printSheet.getRange('B18').setValue(name).setFontFamily('Roboto').setFontSize(10).setFontWeight("bold").setFontColor("#e01b84");
  printSheet.getRange('B19').setValue(companyName).setFontFamily('Roboto').setFontSize(10).setFontColor("#e01b84");
  printSheet.getRange('B20').setValue(streetAddress).setFontFamily('Roboto').setFontSize(10).setFontColor("#e01b84");
  printSheet.getRange('B21').setValue(city +', ' + state + ' ' + zip).setFontFamily('Roboto').setFontSize(10).setFontColor("#e01b84");
  printSheet.getRange('B22').setValue(phone_number).setFontFamily('Roboto').setFontSize(10).setFontColor("#e01b84");
  printSheet.getRange('B23').setValue(email).setFontFamily('Roboto').setFontSize(10).setFontColor("#e01b84");
  
  printSheet.getRange('B11').setValue(date).setFontFamily('Roboto').setFontSize(10).setFontColor("#e01b84");
  printSheet.getRange('D11').setValue(invoice_number).setFontFamily('Roboto').setFontSize(10).setFontColor("#e01b84");
  printSheet.getRange('F11').setValue(po_number).setFontFamily('Roboto').setFontSize(10).setFontColor("#e01b84");
  printSheet.getRange('B14').setValue(ship_date).setFontFamily('Roboto').setFontSize(10).setFontColor("#e01b84");
  printSheet.getRange('D14').setValue(ship_via).setFontFamily('Roboto').setFontSize(10).setFontColor("3e01b84");
  printSheet.getRange('F14').setValue(terms).setFontFamily('Roboto').setFontSize(10).setFontColor("3e01b84");
  
  printSheet.getRange('H28').setValue(ship_and_handle).setFontFamily('Roboto').setFontSize(10).setFontColor("#E01B84");
  printSheet.getRange('H29').setValue(tax_rate).setFontFamily('Roboto').setFontSize(10).setFontColor("#e01b84");
  
  
  // GET LAST ROW OF PO SHEET
  var lastrowPO = poSheet.getLastRow();

  var z = 0;
  var subTotal = 0;
  for(var y = 17; y <= lastrowPO; y++)
  {
    //INSERT ROW ON PRINT SHEET
    printSheet.insertRowsAfter(26 + z, 1);
    
    //GET ITEM VALUES FROM PO SHEET
    var part = poSheet.getRange(y, 1).getValue();
    var description = poSheet.getRange(y, 2).getValue();
    var quantity = poSheet.getRange(y, 3).getValue();
    var unitPrice = poSheet.getRange(y, 4).getValue();
    
    // PRICE TOTALS
    var totalPrice = quantity * unitPrice;
    subTotal = subTotal + totalPrice;
    
    // POPULATE TOTALS ON PRINT SHEET
    printSheet.getRange(26 + z + 1, 2).setValue(part).setFontFamily('Roboto').setFontSize(10).setFontColor("black");
    printSheet.getRange(26 + z + 1, 3).setValue(description).setFontFamily('Roboto').setFontSize(10).setFontColor("#e01b84");
    printSheet.getRange(26 + z + 1, 6).setValue(quantity).setFontFamily('Roboto').setFontSize(10).setFontColor("#e01b84");
    printSheet.getRange(26 + z + 1, 7).setValue(unitPrice).setFontFamily('Roboto').setFontSize(10).setFontColor("#e01b84");
    printSheet.getRange(26 + z + 1, 8).setValue(totalPrice).setFontFamily('Roboto').setFontSize(10).setFontColor("#e01b84");
    
    z++;
  }
  
  // SET TOTAL
  printSheet.getRange(26 + z + 1, 8)
  .setValue(subTotal)
  .setNumberFormat("$#,###.00")
  .setFontFamily('Roboto')
  .setFontSize(10)
  .setFontColor("black");
  
  var totalPO = subTotal;
  
  // CALL PO LOG
  POLog(po_number, name, date, ship_date, totalPO)

}

function POLog(po_number, name, date, ship_date, totalPO)
{
  
   //DEFINE ALL ACTIVE SHEETS
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  //DEFINE PO LOG SHEET          
  var POLogSheet = ss.getSheetByName("PO LOG"); 
  
  //GET LAST ROW OF PO LOG SHEET
  var nextRowPO = POLogSheet.getLastRow() + 1;
  
  //POPULATE INVOICE LOG
  POLogSheet.getRange(nextRowPO, 1).setValue(po_number);
  POLogSheet.getRange(nextRowPO, 2).setValue(name);
  POLogSheet.getRange(nextRowPO, 3).setValue(date);
  POLogSheet.getRange(nextRowPO, 4).setValue(ship_date);
  POLogSheet.getRange(nextRowPO, 5).setValue(totalPO);

}

function ClearInvoice()
{
    //DEFINE ALL ACTIVE SHEETS
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  //DEFINE PO SHEET          
  var poSheet = ss.getSheetByName("POS");
  
  
  //SET VALUES TO NOTHING
  poSheet.getRange(6,2).setValue("");
  poSheet.getRange(7,2).setValue("");
  poSheet.getRange(8,2).setValue("");  
  poSheet.getRange(9,2).setValue("");  
  poSheet.getRange(10,2).setValue("");
  poSheet.getRange(11,2).setValue("");  
  poSheet.getRange(13,2).setValue("");  
  poSheet.getRange(14,2).setValue("");
  
  //CLEAR ITEMS
  poSheet.getRange("A17:D1000").clear();

}
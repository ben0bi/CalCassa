REM  *****  BASIC  *****

Option Explicit
'sheets
Global Const inventory_sheet = 0
Global Const transaction_sheet = 1
' the range of the inventory
Global Const inventory_range ="A2:A4"
' how many positions to add to the item index (A2-1 = A1 = 1)
Global Const inventory_addy = 1

' inventory columns
Global Const inv_col_name = 0
Global Const inv_col_price = 1
Global Const inv_col_amount = 2

' transaction columns
Global Const trans_col_date = 0
Global Const trans_col_name = 1
Global Const trans_col_amount = 2
Global Const trans_col_ausgaben = 3
Global Const trans_col_einnahmen = 4
Global Const trans_col_sum = 5
Global Const SumSum = "=SUMME(E2-D2)"
Global Const trans_col_total = 6
Global Const TotalSum1 = "=SUMME(F2)"
Global Const TotalSum2 = "=SUMME(G2+F3)"

Public oBuySellDialog

Private isTextChanging
Private actualItemIndex

' count of inventory items (see range: A4-A2 = 2(+1)=3)
Private inventory_count 

' create dialog box for cassa.
Sub START
	Dim oList, oData, i, n

	actualItemIndex = -1
	isTextChanging = 0
	oBuySellDialog = CreateUnoDialog(DialogLibraries.Standard.KassenDialog)
    oList = oBuySellDialog.getControl("combo_inventorylist")

	oList.Text = "Select inventory item."

    'Read the data list from cell range into a variant array '
    oData = ThisComponent.Sheets(inventory_sheet).getCellRangeByName(inventory_range).DataArray

    inventory_count = ubound(oData)

    For i = 0 to inventory_count
      oList.addItem(oData(i)(0) ,i)
    Next i
  
    oBuySellDialog.Execute()
End Sub

' get the index of combo selected text
Function getComboIndex(list)
	Dim txt,i, ret
	ret = -1
	txt = list.getSelectedText
	for i=0 to list.getItemCount-1
		if list.getItem(i) = txt then
			ret = i
		endif
	next
	getComboIndex = ret
end function

' inventory list item has changed.
Sub onListItemClick
	Dim oList, oAmount, i, oCellPrice
	Dim oPrice, oMulti, oTotal 
    oList = oBuySellDialog.getControl("combo_inventorylist")
    
    oPrice = oBuySellDialog.getControl("txt_price")
    oMulti = oBuySellDialog.getControl("txt_multiplier")
    oTotal = oBuySellDialog.getControl("txt_total")
    
    oAmount = oBuySellDialog.getControl("lbl_amount")
 	actualItemIndex = getComboIndex(oList)
    if actualItemIndex>-1 then
	    oCellPrice = ThisComponent.Sheets(inventory_sheet).getCellByPosition(inv_col_price,actualItemIndex+inventory_addy)
    	oPrice.Text = oCellPrice.String
    	oAmount.Text = " Vorhanden: "+ThisComponent.Sheets(inventory_sheet).getCellByPosition(inv_col_amount, actualItemIndex+inventory_addy).String
    endif
    
    ' set total text
    oTotal.Text = CStr(CDbl(oPrice.Text)*CDbl(oMulti.Text))
End Sub

' set value of total when multiplier or price text has changed.
Sub onMultiplierTextChange
	Dim oPrice, oMulti, oTotal
	if isTextChanging = 0 then
		isTextChanging = 1
	    oPrice = oBuySellDialog.getControl("txt_price")
    	oMulti = oBuySellDialog.getControl("txt_multiplier")
	    oTotal = oBuySellDialog.getControl("txt_total")
    	oTotal.Text = CStr(CDbl(oPrice.Text)*CDbl(oMulti.Text))
    	isTextChanging = 0
    endif
End Sub

' set value of price when total text has changed
Sub onTotalTextChange
	Dim oPrice, oMulti, oTotal
	if isTextChanging = 0 then
		isTextChanging = 1
	    oPrice = oBuySellDialog.getControl("txt_price")
    	oMulti = oBuySellDialog.getControl("txt_multiplier")
	    oTotal = oBuySellDialog.getControl("txt_total")
	    if CDbl(oMulti.Text)>0 and CDbl(oTotal.Text)>0 then
	    	oPrice.Text = CStr(CDbl(oTotal.Text)/CDbl(oMulti.Text))
	    else
	    	oPrice.Text="0"
	    endif
    	isTextChanging = 0
    endif
End Sub

' the buy button was clicked.
Sub onBuyBtnClicked
	Dim oList, oAmount, oCell,oCellName, price, amt
	Dim oPrice, oMulti
'	MsgBox(actualItemIndex)
    oPrice = oBuySellDialog.getControl("txt_price")
   	oMulti = oBuySellDialog.getControl("txt_multiplier")
	oAmount = oBuySellDialog.getControl("lbl_amount")	
    oList = oBuySellDialog.getControl("combo_inventorylist")
	
	price = CDbl(oPrice.Text)*CDbl(oMulti.Text)

	if actualItemIndex>-1 and CDbl(oMulti.Text)>0 then
		oCellName = ThisComponent.Sheets(inventory_sheet).getCellByPosition(inv_col_name, actualItemIndex+inventory_addy)
		oCell = ThisComponent.Sheets(inventory_sheet).getCellByPosition(inv_col_amount, actualItemIndex+inventory_addy)
		amt = oCell.String
		amt = CDbl(amt) + CDbl(oMulti.Text)
		oCell.String = amt
    	oAmount.Text = " Vorhanden: "+oCell.String
    	MsgBox("EINGEKAUFT: "+oMulti.Text+" * "+oCellName.String+" für "+price+" CHF")
	endif
 	createTransaction("[E]"+oMulti.Text+" * "+oList.Text, CDbl(oMulti.Text), price, 0)
'	createTransaction("E", CDbl(oMulti.Text), price, 0)
End Sub


' the sell button was clicked.
Sub onSellBtnClicked
	Dim oList, oAmount, oCell,oCellName, price, amt
	Dim oPrice, oMulti
'	MsgBox(actualItemIndex)
    oPrice = oBuySellDialog.getControl("txt_price")
   	oMulti = oBuySellDialog.getControl("txt_multiplier")
	oAmount = oBuySellDialog.getControl("lbl_amount")	
    oList = oBuySellDialog.getControl("combo_inventorylist")

	price = CDbl(oPrice.Text)*CDbl(oMulti.Text)
	if CDbl(oMulti.Text)>0 then
		if actualItemIndex>-1 then
			oCellName = ThisComponent.Sheets(inventory_sheet).getCellByPosition(inv_col_name, actualItemIndex+inventory_addy)
			oCell = ThisComponent.Sheets(inventory_sheet).getCellByPosition(inv_col_amount, actualItemIndex+inventory_addy)
			amt = oCell.String
			amt = CDbl(amt) - CDbl(oMulti.Text)
			if CDbl(amt) <0 then 
				amt = 0
			endif
			oCell.String = amt
    		oAmount.Text = " Vorhanden: "+oCell.String
    		MsgBox("[i] VERKAUFT: "+oMulti.Text+" * "+oCellName.String+" fuer "+price+" CHF")
    		createTransaction(CStr(oMulti.Text)+" * "+oCellName.String, CDbl(oMulti.Text), 0, price)
    	else
    		MsgBox("[n] VERKAUFT: "+oMulti.Text+" * "+oList.Text+" fuer "+price+" CHF")
    		createTransaction(CStr(oMulti.Text)+" * "+oList.Text, CDbl(oMulti.Text), 0, price)
    	endif
	endif
End Sub

Sub createTransaction(thename as String, amount as Double, ausgaben as Double, einnahmen as Double)
	Dim oSheet, oCell
	oSheet = ThisComponent.Sheets(transaction_sheet)
	oSheet.Rows.insertByIndex(1,1)
	
	' TODO: SET DATE
	
	' set name
	oCell = oSheet.getCellByPosition(trans_col_name,1)
	oCell.String = thename
	' set amount
	oCell = oSheet.getCellByPosition(trans_col_amount,1)
	oCell.String = amount
	' set ausgaben
	oCell = oSheet.getCellByPosition(trans_col_ausgaben,1)
	oCell.String = ausgaben
	' set einnahmen
	oCell = oSheet.getCellByPosition(trans_col_einnahmen,1)
	oCell.String = einnahmen
	' set sums
	oCell = oSheet.getCellByPosition(trans_col_sum,1)
	oCell.String = SumSum
	oCell = oSheet.getCellByPosition(trans_col_total,1)
	oCell.String = TotalSum1
	oCell = oSheet.getCellByPosition(trans_col_total,2)
	oCell.String = TotalSum2
End Sub


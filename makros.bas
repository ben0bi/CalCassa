REM  *****  BASIC  *****

Option Explicit
Global Const inventory_sheet = 0
Global Const inventory_range ="A2:A4"
Global Const inventory_count = 3
Global Const inventory_addy = 1

Public oBuySellDialog

Private isTextChanging

' create dialog box for cassa.
Sub START
	Dim oList, oData, i, n

	isTextChanging = 0
	oBuySellDialog = CreateUnoDialog(DialogLibraries.Standard.KassenDialog)
    oList = oBuySellDialog.getControl("combo_inventorylist")

	oList.Text = "Select inventory item."

    'Read the data list from cell range into a variant array '
    oData = ThisComponent.Sheets(inventory_sheet).getCellRangeByName(inventory_range).DataArray

    n = ubound(oData)

    For i = 0 to n
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
	Dim oList, oAmount, item, i, oCellPrice
	Dim oPrice, oMulti, oTotal 
    oList = oBuySellDialog.getControl("combo_inventorylist")
    
    oPrice = oBuySellDialog.getControl("txt_price")
    oMulti = oBuySellDialog.getControl("txt_multiplier")
    oTotal = oBuySellDialog.getControl("txt_total")
    
    oAmount = oBuySellDialog.getControl("lbl_amount")
 	item = getComboIndex(oList)
    if item>-1 then
	    oCellPrice = ThisComponent.Sheets(inventory_sheet).getCellByPosition(1,item+inventory_addy)
    	oPrice.Text = oCellPrice.String
    	oAmount.Text = " Vorhanden: "+ThisComponent.Sheets(inventory_sheet).getCellByPosition(2, item+inventory_addy).String
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


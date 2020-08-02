REM  *****  BASIC  *****

Sub Main
	Dim price as Long
	Dim priceA as Integer
	Dim priceB as Integer
	Dim temp as Integer
	
	Dim currDoc as Object
	Dim currSheet as Object
	Dim currCell as Object
	
	currDoc = ThisComponent
	currSheet = currDoc.sheets(0)
	
	Const casePrice as Integer = 2000
	price = 0
	temp = 0
	
	For i = 0 To 100 Step 1
		currCell = currSheet.getCellByPosition(5, i)
		
		
		If currCell.String <> "" Then
			priceA = currSheet.getCellByPosition(1, i).Value
			priceB = currSheet.getCellByPosition(2, i).Value
			
			If priceA = 0 Then
				temp = priceB
			ElseIf priceB = 0 Then
				temp = priceA
			ElseIf priceA < priceB Then
				temp = priceA
			Else
				temp = priceB
			End If
			
			price = price + temp
		End If
	Next
	
	price = price + casePrice
	currCell = currSheet.getCellByPosition(6, 1)
	currCell.Value = price
	currCell = currSheet.getCellByPosition(7, 1)
	currCell.Value = price \ 61
	
End Sub

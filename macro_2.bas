REM  *****  BASIC  *****
Global currDoc as Object
Global currSheet as Object
Global currCell as Object


Sub Main
	currDoc = ThisComponent
	currSheet = currDoc.sheets(0)
	
	Dim price as Long
	Dim priceA as Integer
	Dim priceB as Integer
	Dim temp as Integer

	Dim itemsA() as Variant
	Dim itemsB() as Variant
	Dim Ai as Integer
	Dim Bi as Integer
	
	Const casePrice as Integer = 2000
	
	price = 0
	temp = 0
	Ai = 0
	Bi = 0
	
	For i = 0 To GetRange("F1") Step 1
		currCell = currSheet.getCellByPosition(5, i)
		
		
		If currCell.String <> "" Then
			priceA = currSheet.getCellByPosition(1, i).Value
			priceB = currSheet.getCellByPosition(2, i).Value
			
			temp = ReturnLesser(priceA, priceB)
			If temp = priceA Then
				ReDim Preserve itemsA(Ai)
				itemsA(Ai) = currSheet.getCellByPosition(0, i).String
				Ai = Ai + 1
			Else
				ReDim Preserve itemsB(Bi)
				itemsB(Bi) = currSheet.getCellByPosition(0, i).String
				Bi = Bi + 1
			End If
			
			price = price + temp
		End If
	Next
	
	
	price = price + casePrice
	
	Call WriteArray(itemsA, itemsB)
	Call SetPrice(price)
End Sub

Function GetRange(cellName as Variant) as Integer
	Dim Cur as Object
	Dim Range as Object
	Cur = currSheet.createCursorByRange(currSheet.getCellRangeByName(cellName))
	Cur.gotoEndOfUsedArea(True)
	Range = currSheet.getCellRangeByName(Cur.AbsoluteName)
	GetRange = Range.RangeAddress.EndRow
End Function

Function ReturnLesser(num1 as Integer, num2 as Integer) as Integer
	If num1 = 0 Then
		ReturnLesser = num2
	ElseIf num2 = 0 Then
		ReturnLesser = num1
	ElseIf num1 < num2 Then
		ReturnLesser = num1
	Else
		ReturnLesser = num2
	End If
End Function

Sub SetPrice(price)
	currCell = currSheet.getCellByPosition(6, 1)
	currCell.Value = price
	currCell = currSheet.getCellByPosition(7, 1)
	currCell.Value = price \ 61
End Sub

Sub WriteArray(itemsA as Variant, itemsB as Variant)
	For i = 0 To ArrayLen(itemsA) - 1 Step 1
		currCell = currSheet.getCellByPosition(8, i + 1)
		currCell.String = itemsA(i)
	Next
	
	For i = 0 To ArrayLen(itemsB) - 1 Step 1
		currCell = currSheet.getCellByPosition(9, i + 1)
		currCell.String = itemsB(i)
	Next
End Sub

Function ArrayLen(arr as Variant) as Integer
	ArrayLen = UBound(arr) - LBound(arr) + 1
End Function
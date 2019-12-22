Sub ConsolidateStockData()
	For Each ws In Worksheets
		Call ConsolidateStockDataForWS(ws)
	Next ws
End Sub


Sub ConsolidateStockDataForWS(currentSheet)
	lastRow = currentSheet.Cells(Rows.Count, "A").End(xlUp).Row
	outRowCount = 2
	
	currentSticker = currentSheet.Range("A2").Value
	currentOpenPrice = currentSheet.Range("C2").Value
	currentVolume = currentSheet.Range("G2").Value

	currentSheet.Range("I" & 1).Value = "Ticker"
	currentSheet.Range("J" & 1).Value = "Yearly Change"
	currentSheet.Range("K" & 1).Value = "Percent Change"
	currentSheet.Range("L" & 1).Value = "Total Stock Volume"
	
	
	For i = 2 To lastRow
		newSticker = currentSheet.Range("A" & i).Value
		If currentSticker = newSticker Then
			currentVolume = currentVolume + currentSheet.Range("G" & i).Value
		Else
			currentClosePrice = currentSheet.Range("F" & i).Value
			priceChange = currentClosePrice - currentOpenPrice
			If currentOpenPrice > 0 then
				percentChange = priceChange / currentOpenPrice
			End if
			
			currentSheet.Range("I" & outRowCount).Value = currentSticker
			currentSheet.Range("J" & outRowCount).Value = priceChange
			currentSheet.Range("K" & outRowCount).Value = percentChange
			currentSheet.Range("K" & outRowCount).Style = "Percent"
			currentSheet.Range("K" & outRowCount).NumberFormat = "0.00%"
			currentSheet.Range("L" & outRowCount).Value = currentVolume

			If percentChange < 0 then
				currentSheet.Range("K" & outRowCount).Interior.ColorIndex = 3
			Else
				currentSheet.Range("K" & outRowCount).Interior.ColorIndex = 4
			End If

			' set for next loop
			currentSticker = newSticker
			currentOpenPrice = currentSheet.Range("C" & i).Value
			currentVolume = currentSheet.Range("G" & i).Value
			outRowCount = outRowCount + 1
		End If
	Next
	
End Sub

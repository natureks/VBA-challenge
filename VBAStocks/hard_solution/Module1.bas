Sub ConsolidateStockData()
	 For Each ws In Worksheets
		  Call ConsolidateStockDataForWS(ws)
	 Next ws
End Sub


Sub ConsolidateStockDataForWS(currentSheet)
	 lastRow = currentSheet.Cells(Rows.Count, "A").End(xlUp).Row
	 outRowCount = 2
	 
	greatestPercentIncrease = 0
	greatestPercentDecrease = 0
	greatestVolume = 0

	greatestPercentIncreaseTicker = ""
	greatestPercentDecreaseTicker = ""
	greatestVolumeTicker = ""

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

			' for greatest stats
			If percentChange < 0 Then
				If percentChange < greatestPercentDecrease Then
					greatestPercentDecrease = percentChange
					greatestPercentDecreaseTicker = currentSticker
				End If
			ElseIf percentChange > greatestPercentIncrease Then
				greatestPercentIncrease = percentChange
				greatestPercentIncreaseTicker = currentSticker
			End If

			If currentVolume > greatestVolume Then
				greatestVolume = currentVolume
				greatestVolumeTicker = currentSticker
			End If
			
			' set for next loop
			currentSticker = newSticker
			currentOpenPrice = currentSheet.Range("C" & i).Value
			currentVolume = currentSheet.Range("G" & i).Value
			outRowCount = outRowCount + 1
		  End If
	 Next
	 
	currentSheet.Range("O" & 2).Value = "Greatest % Increase"
	currentSheet.Range("O" & 3).Value = "Greatest % Decrease"
	currentSheet.Range("O" & 4).Value = "Greatest Total Volume"

	currentSheet.Range("P" & 1).Value = "Ticker"
	currentSheet.Range("P" & 2).Value = greatestPercentIncreaseTicker
	currentSheet.Range("P" & 3).Value = greatestPercentDecreaseTicker
	currentSheet.Range("P" & 4).Value = greatestVolumeTicker

	currentSheet.Range("Q" & 1).Value = "Value"
	currentSheet.Range("Q" & 2).Value = greatestPercentIncrease
	currentSheet.Range("Q" & 2).Style = "Percent"
	currentSheet.Range("Q" & 2).NumberFormat = "0.00%"
	currentSheet.Range("Q" & 3).Value = greatestPercentDecrease
	currentSheet.Range("Q" & 3).Style = "Percent"
	currentSheet.Range("Q" & 3).NumberFormat = "0.00%"
	currentSheet.Range("Q" & 4).Value = greatestVolume
End Sub

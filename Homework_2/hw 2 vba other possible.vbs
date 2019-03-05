sub secondidea():

dim ws as worksheet
	for each ws in activeworkbook.worksheets
	ws.activate

	dim i as long
	dim ticker as string
	dim volume as double 
	dim column as integer
	dim row as double
	dim close_pr as double
	dim open_pr as double
	dim YRchange as double
	dim PRchange as double

	row = 2
	column = 1
	volume = 0
	lastrow = Cells(Rows.Count, 1).End(xlUp).Row
	open_pr = cells(2, column + 2).Value

		for i = 2 to lastrow
			
			if cells(i + 1, column).Value <> cells(i, column).Value then 
				'ticker name '
				ticker = cells(i, column).Value
				'add last row of same ticker volume to total'
				volume = volume + cells(i, column + 6).Value
				'post ticker name'
				cells(row, column + 9).Value = ticker
				'post volume total'
				cells(row, column + 12).Value = volume
				'get close price'
				close_pr = cells(i, column + 5).Value
				'calc yearly change'
				YRchange = close_pr - open_pr 
				'post yearly change'
				cells(row, column + 10).Value = YRchange
				
				if open_pr = 0 and close_pr = 0 then
					PRchange = 0
				elseif open_pr = 0 and close_pr <> 0 then
					PRchange = 1
				else 
					'calc percent change'
					PRchange = YRchange / open_pr
					'post percent change'
					cells(row, column + 11).Value = PRchange
				end if 
					'color added'
					if YRchange >= 0 then 
						cells(row, column + 10).interior.colorindex = 4
					elseif YRchange < 0 then 
						cells(row, column + 10).interior.colorindex = 3 
					end if 	

				'restart open price'
				open_pr = cells(i + 1, column + 2).Value 
				'post next ticker in following line'
				row = row + 1
				'reset the total volume'
				volume = 0

			else 
				'cont. to add volume until ticker change'
				volume = volume + cells(i, column + 6).Value

			end if

		next i

	'headings for the data summary'
	range("J1").Value = "Ticker"
	range("K1").Value = "Yearly Change"
	range("L1").Value = "Percent Change"
	range("M1").Value = "Total Stock Volume"
	'final adjustments'
	columns("J:M").AutoFit
	columns("L").numberformat = "0.00%"


	next ws 
end sub
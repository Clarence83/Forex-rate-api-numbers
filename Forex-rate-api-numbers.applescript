set arrList to {}
set currencyList to {}
set baseSymbol to {"USD"}

set startDate to short date string of ((current date) - 1 * days)
set endDate to short date string of (current date)

tell application "Finder"
	-- define the path of all handler addresses
	set mypath to ((container of (path to me) as string) & "Function:functions.scpt") as alias
	set myKey to ((container of (path to me) as string) & "Function:apiKey.scpt") as alias
	
	-- define the path of numbers spreadsheet
	set chosenDocumentFile to POSIX path of (container of (path to me) as alias) & "Forexprices.numbers"
end tell

set retrieveapiKey to load script myKey
set myapiKey to apiKey() of retrieveapiKey
set symbol to symbol() of retrieveapiKey
-- convert to string and separate each symbol with a commas
set AppleScript's text item delimiters to ","
set symbols to symbol as string

set myOtherScript to load script mypath
set loadScript to currentTime(startDate, endDate) of myOtherScript
set {start_Date, end_Date} to the result

set loadScript to JSONhelper(baseSymbol, myapiKey, start_Date, end_Date, symbols) of myOtherScript
set arrList to the result

tell application "Numbers"
	try
		open the chosenDocumentFile
		
		try
			
			if not (exists document 1) then error number 1000
			tell document 1
				set selectedTable to table 1 of sheet "Forex Rates"
				try
				on error
					error number 1001
				end try
				
				tell selectedTable
					set cntRow to count row
					set cntCol to count column
					
					set curRow to 2
					
					repeat with i from 1 to length of symbol
						--column 1 setup
						set value of cell curRow of column 1 to item i of symbol & "/USD"
						
						if currencyList contains "SGD" then
							return item i of symbol as text
							
						end if
						set curRow to curRow + 1
						-- add new row to bottom when current row is greater than count row
						if cntRow < curRow then
							-- store the address of the current non-header cell area
							set tableRangeAddress to (name of cell 2 of row 2) & ":" & (name of last cell)
							-- add a row to the end of the table
							add row below last row
							set cntRow to cntRow + 1
						end if
					end repeat
					
					set timeConvert to (date "Thursday, 1 January 1970 at 8:00:00 AM") + (timestamp of item 2 of arrList) / 86400 * days
					
					set curRow to 2
					
					set newsymbol to {}
					set newInfo to {}
					set currencyRate to rates of item 2 of arrList
					set currencyInfo to rates of item 1 of arrList
					
					-- in order to get array of newsymbol
					set SGD to SGD of currencyRate
					set USD to USD of currencyRate
					set EUR to EUR of currencyRate
					set JPY to JPY of currencyRate
					set GBP to GBP of currencyRate
					set AUD to AUD of currencyRate
					set CAD to CAD of currencyRate
					set CHF to CHF of currencyRate
					set CNY to CNY of currencyRate
					set NZD to NZD of currencyRate
					set MYR to MYR of currencyRate
					set INR to INR of currencyRate
					
					set newsymbol to newsymbol & SGD & USD & EUR & JPY & GBP & AUD & CAD & CHF & CNY & NZD & MYR & INR
					
					repeat with i from 1 to length of newsymbol
						set value of cell curRow of column 2 to item i of newsymbol
						set value of cell curRow of column 3 to timeConvert
						set curRow to curRow + 1
					end repeat
					
					set curRow to 2
					
					set value of cell curRow of column 4 to start_rate of SGD of currencyInfo
					set value of cell curRow of column 5 to end_rate of SGD of currencyInfo
					set value of cell curRow of column 6 to change of SGD of currencyInfo
					set value of cell curRow of column 7 to change_pct of SGD of currencyInfo
					
					set curRow to curRow + 1
					set value of cell curRow of column 4 to start_rate of USD of currencyInfo
					set value of cell curRow of column 5 to end_rate of USD of currencyInfo
					set value of cell curRow of column 6 to change of USD of currencyInfo
					set value of cell curRow of column 7 to change_pct of USD of currencyInfo
					
					set curRow to curRow + 1
					set value of cell curRow of column 4 to start_rate of EUR of currencyInfo
					set value of cell curRow of column 5 to end_rate of EUR of currencyInfo
					set value of cell curRow of column 6 to change of EUR of currencyInfo
					set value of cell curRow of column 7 to change_pct of EUR of currencyInfo
					
					set curRow to curRow + 1
					set value of cell curRow of column 4 to start_rate of JPY of currencyInfo
					set value of cell curRow of column 5 to end_rate of JPY of currencyInfo
					set value of cell curRow of column 6 to change of JPY of currencyInfo
					set value of cell curRow of column 7 to change_pct of JPY of currencyInfo
					
					set curRow to curRow + 1
					set value of cell curRow of column 4 to start_rate of GBP of currencyInfo
					set value of cell curRow of column 5 to end_rate of GBP of currencyInfo
					set value of cell curRow of column 6 to change of GBP of currencyInfo
					set value of cell curRow of column 7 to change_pct of GBP of currencyInfo
					
					set curRow to curRow + 1
					set value of cell curRow of column 4 to start_rate of AUD of currencyInfo
					set value of cell curRow of column 5 to end_rate of AUD of currencyInfo
					set value of cell curRow of column 6 to change of AUD of currencyInfo
					set value of cell curRow of column 7 to change_pct of AUD of currencyInfo
					
					set curRow to curRow + 1
					set value of cell curRow of column 4 to start_rate of CAD of currencyInfo
					set value of cell curRow of column 5 to end_rate of CAD of currencyInfo
					set value of cell curRow of column 6 to change of CAD of currencyInfo
					set value of cell curRow of column 7 to change_pct of CAD of currencyInfo
					
					set curRow to curRow + 1
					set value of cell curRow of column 4 to start_rate of CHF of currencyInfo
					set value of cell curRow of column 5 to end_rate of CHF of currencyInfo
					set value of cell curRow of column 6 to change of CHF of currencyInfo
					set value of cell curRow of column 7 to change_pct of CHF of currencyInfo
					
					set curRow to curRow + 1
					set value of cell curRow of column 4 to start_rate of CNY of currencyInfo
					set value of cell curRow of column 5 to end_rate of CNY of currencyInfo
					set value of cell curRow of column 6 to change of CNY of currencyInfo
					set value of cell curRow of column 7 to change_pct of CNY of currencyInfo
					
					set curRow to curRow + 1
					set value of cell curRow of column 4 to start_rate of NZD of currencyInfo
					set value of cell curRow of column 5 to end_rate of NZD of currencyInfo
					set value of cell curRow of column 6 to change of NZD of currencyInfo
					set value of cell curRow of column 7 to change_pct of NZD of currencyInfo
					
					set curRow to curRow + 1
					set value of cell curRow of column 4 to start_rate of MYR of currencyInfo
					set value of cell curRow of column 5 to end_rate of MYR of currencyInfo
					set value of cell curRow of column 6 to change of MYR of currencyInfo
					set value of cell curRow of column 7 to change_pct of MYR of currencyInfo
					
					set curRow to curRow + 1
					set value of cell curRow of column 4 to start_rate of INR of currencyInfo
					set value of cell curRow of column 5 to end_rate of INR of currencyInfo
					set value of cell curRow of column 6 to change of INR of currencyInfo
					set value of cell curRow of column 7 to change_pct of INR of currencyInfo
					
				end tell
			end tell
			
		on error errorMessage number errorNumber
			if errorNumber is 1000 then
				set alertString to "MISSING RESOURCE"
				set errorMessage to "Please create or open a document before running this script."
			else if errorNumber is 1001 then
				set alertString to "SELECTION ERROR"
				set errorMessage to "Please select a table before running this script."
			else
				set alertString to "EXECUTION ERROR"
			end if
			display alert alertString message errorMessage buttons {"Cancel"}
			error number -128
		end try
	end try
end tell

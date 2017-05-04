# ExcelToCal.ps1
# Creates a CSV from a workbook and imports it to Outlook

# Create CSV

# Creating excel object, set it to run headless, ignore any alerts (default to "Yes" response)
$Excel = New-Object -ComObject Excel.Application
$Excel.visible = $False
$Excel.DisplayAlerts = $False

# Create a Workbook excel object, point to the correct workbook, and specify a sheet
$Workbook = $Excel.Workbooks.Open("[WORKBOOK LOCATION\Name]")
$Worksheet = $Workbook.sheets.item("[SHEET NAME]")

# Verify the workbook is open and display active sheets
$Workbook.sheets | Select-Object -Property Name

# Specify CSV name and location, and to save as a csv (6 = CSV)
$Workbook.SaveAs("[CSV LOCATION\NAME.csv]", 6)

# Close the excel file and kill the process
$excel.quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)

# End Create CSV

# Import to Cal

# Create Outlook object
$Outlook = New-Object -ComObject Outlook.Application
$OutlookNamespace = $Outlook.GetNamespace("MAPI")

# Set the default Outlook calendar and csv file
$OutlookCalendar = $OutlookNamespace.GetDefaultFolder(9)
$SourceCSV = Import-CSV "[CSV LOCATION\NAME.csv]"

# Create a list the same length as the CSV file
$listCount = $SourceCSV.count

# Create an array from the CSV, iterate through it adding each line item to the Calendar 
#The part after $SourceCSV[$i] is the title of the column you wish to grab the values from
for($i = 0; $i -lt $listCount; ++$i)
{
	$CalItem = $Outlook.CreateItem(1)
	$CalItem.Subject = $SourceCSV[$i].subject
	$CalItem.start = $SourceCSV[$i].start
	$CalItem.end = $SourceCSV[$i].end
	$CalItem.AllDayEvent = $True
	# $CalItem.Body = $SourceCSV[$i].body
	# $CalItem.Location = $SourceCSV[$i].location
	# $CalItem.Importance = $SourceCSV[$i].importance
	# $CalItem.BusyStatus = $SourceCSV[$i].busyStatus
	# $CalItem.EnableReminder = $SourceCSV[$i].enableReminder
	# $CalItem.MeetingStart = $SourceCSV[$i].meetingStart
	# $CalItem.MeetingDuration = $SourceCSV[$i].meetingDuration
	# $CalItem.Reminder = $SourceCSV[$i].reminder
	$a = $CalItem.save()
}

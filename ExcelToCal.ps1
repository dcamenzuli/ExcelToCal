# ExcelToCal.ps1
# Creates a CSV from a workbook and imports it to Outlook

# Create CSV

# Creating excel object, set it to run headless, ignore any alerts (default to "Yes" response)
$Excel = New-Object -ComObject Excel.Application
$Excel.visible = $False
$Excel.DisplayAlerts = $False

# Create a Workbook excel object, point to the correct workbook, and specify a sheet
$Workbook = $Excel.Workbooks.Open("[CSVLocation\Name]")
$Worksheet = $Workbook.sheets.item("[SheetName]")

# Verify the workbook is open and display active sheets
$Workbook.sheets | Select-Object -Property Name

# Specify CSV name and location, and to save as a csv (6 = CSV)
$Workbook.SaveAs("[CSVLocation\Name.csv]", 6)

# Close the excel file and kill the process
$excel.quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)

# End Create CSV


# Import to Cal

# Create Outlook object
$Outlook = New-Object -ComObject Outlook.Application
$OutlookNamespace = $Outlook.GetNamespace("MAPI")

# Set the default Outlook calendar to the shared Calendar
$myRecipient = $OutlookNamespace.CreateRecipient("[SharedCalendarName]")

If($myRecipient.Resolve())
{
	$OutlookCalendar = $OutlookNamespace.GetSharedDefaultFolder($myRecipient, 9)
}

# Set Source CSV
$SourceCSV = Import-CSV "[CSVLocation\Name.csv]"

# Create an integer to represent the size of the CSV
$listCount = $SourceCSV.Count

# Create an array from the CSV, iterate through it adding each line item to the Calendar 
# The part after $SourceCSV[$i] is the title of the column you wish to grab the values from
for($i = 0; $i -lt $listCount; ++$i)
{
	$CalItem = $OutlookCalendar.Items.Add(1)
	$CalItem.Subject = $SourceCSV[$i].subject
	$CalItem.start = $SourceCSV[$i].start
	$CalItem.AllDayEvent = $True
	$CalItem.ReminderSet = $False
	# $CalItem.end = $SourceCSV[$i].end
	# $CalItem.Body = $SourceCSV[$i].body
	# $CalItem.Location = $SourceCSV[$i].location
	# $CalItem.Importance = $SourceCSV[$i].importance
	# $CalItem.BusyStatus = $SourceCSV[$i].busyStatus
	# $CalItem.MeetingStart = $SourceCSV[$i].meetingStart
	# $CalItem.MeetingDuration = $SourceCSV[$i].meetingDuration
	# $CalItem.Reminder = $SourceCSV[$i].reminder
	
	# Check if entry exists before adding new entry
	$checkInt = 0
	foreach($b in $OutlookCalendar.items)
	{
		if($b.Subject -eq $CalItem.Subject -and $b.Start -eq $CalItem.Start)
		{
			++$checkInt
		}
	}
	if($checkInt -eq 0)
	{
		$a = $CalItem.save()
	}
}
$Outlook.quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($Outlook)

# End Import to Cal

Some power shell scripts

##########################################################

Scipt to local copy

# Code to copy file to google drive
# User should have enebaled google drive in gmail account
# User need to have google drive installed on this machine
# User should be able to acces the google drive folder locally using command prompt 

function Local_Copy([String] $source , [String] $destination)
{

$filePath=$destination
Write-Host "Destination path defined"
Copy-Item -Path $source -Destination $filePath
Write-Host "File copied to destination"
if (!(Test-Path $filePath))
{
Write-Host "File doesn't exist "
throw "File copy operation failed !"
}
Write-Host "File is created .Script end here !!!"
}

Local_Copy "E:\Powershell script\World\Word file creation\MsWord.docx" "C:\Users\deepakt\Google Drive\curtain\"

###########################################################

Scripts for generating log files

# This script is used to generate the log file
# In this script you need to define path for log file
# Function "Logger" accepts message as input
# Function "Logger" generates the date time for each line


$global:LogFilePath = 'E:\Powershell script\World\DELL_Automation.log'

function Logger
{
    param (
        	[Parameter(Mandatory)]
        	[string]$Message
    )
    
    $line = [pscustomobject]@{
       		'DateTime' = (Get-Date)
        	'Message'  = $Message
   }
    $line | Export-Csv -Path $LogFilePath -Append -NoTypeInformation
}

 Logger 'Please check the log file'

#########################################################


Scripts for  remote copy


# This script is used for remote copy of single file
# Make sure that remote path is accessible using UNC path or destination is shared with you
# Make sure that source and destination machine are connected using LAN
# Make sure that source and destination machine are on same network
# Please define your destination machine UNC path


function Remote_Copy([string] $SOURCE ,[string] $DEST ,[string] $filename)

{

try {
		Write-Host 'Defining source directory'
		$source=$SOURCE
		Write-Host $source
		$destination=$DEST
		Write-Host 'Defining destination directory'
		Write-Host $destination
		Write-Host Printing file to copy 
		Write-Host $filename
		Write-Host 'Initializing copy operation........'
		Copy-Item $source$filename -Destination $destination -force 
		Write-Host 'File copied...'
		Write-Host 'File is created at' $destination$filename
		Write-Host 'Function return value...'
		return $destination+$filename
		}
catch	{
		Write-Host $ErrorMessage = $_.Exception.Message
    	Write-Host $FailedItem = $_.Exception.ItemName
		Break 
		}
finally {
		if (!(Test-Path ($destination+$filename)))
		{
		Write-Host 'File not copied'
		Throw 'Copy operation failed *************'
		}
		}
}

Remote_Copy 'E:\Powershell script\World\Word file creation\' '\\ULTP_488\SaGaR_Share\' 'MsWord.docx'

############################################################

Script to fire an email

#This script is used to send mail using smtp server
#From account theautomationtestreport@gmail.com
#To account deepak.tiwari@synerzip.com
#Attachment is enable
#HTML body content is enable
#SSL is enable
#Mail priority is enable [ works in outlook ]


Write-Host "Defining smtp server"
$emailSmtpServer = "smtp.gmail.com"
Write-Host "Defining smtp port"
$emailSmtpServerPort = "587"
Write-Host "Defining credentials"
$emailSmtpUser = "theautomationtestreport@gmail.com"
$emailSmtpPass = "synerzip"
 
$emailMessage = New-Object System.Net.Mail.MailMessage
Write-Host "Defining from account"
$emailMessage.From = "THE POWERSHELL GUY <theautomationtestreport@gmail.com>"
Write-Host "Defining to account"
$emailMessage.To.Add( "deepak.tiwari@synerzip.com" )
#$emailMessage.To.Add( "kumar.deepaktiwari@gmail.com" )
$emailMessage.Subject = "PROTECTED DOC"
$attachment = "E:\Powershell script\World\Word file creation\MsWord.docx"
Write-Host "Constructing mail body with attachment"
$emailMessage.Attachments.Add( $attachment )
Write-Host "Enabling mail send notification on success"
$emailMessage.DeliveryNotificationOptions= [System.Net.Mail.DeliveryNotificationOptions]::OnSuccess # This doesn't work
Write-Host "Enabling mail body with HTML content"
$emailMessage.IsBodyHtml = $true
$emailMessage.Body = @"
<p>Please open the <strong>Protected Document</strong>.</p>
<p>at your end</p>
"@
Write-Host "Enabling mail priority high"
$emailMessage.Priority = [System.Net.Mail.MailPriority]::High
$SMTPClient = New-Object System.Net.Mail.SmtpClient( $emailSmtpServer , $emailSmtpServerPort )
Write-Host "Enabling SSL"
$SMTPClient.EnableSsl = $true
$SMTPClient.Credentials = New-Object System.Net.NetworkCredential( $emailSmtpUser , $emailSmtpPass );
$SMTPClient.Send( $emailMessage )
Write-Host "Mail send completed,Script end here !!!"


################################################################

Script to  create MSword file with content

# This script is used to write MSword file
# This script is used to write content in MSword file


Write-Host "Defining date"
$date = get-date -format MM-dd-yyyy
Write-Host "Defining file path"
$filePath = "E:\Powershell script\World\Word file creation\MsWord.docx"
Write-Host "Checking if file already exist"
if (Test-Path $filePath)
{
Write-Host "File already exist"
Remove-Item $filePath
Write-Host "File removed"
}
[ref]$SaveFormat = "microsoft.office.interop.word.WdSaveFormat" -as [type]
$word = New-Object -ComObject word.application
#Write-Host "Setting visibility of file"
#$word.visible = $true
$doc = $word.documents.add()
Write-Host "Making selection in file"
$selection = $word.selection
$selection.WholeStory
Write-Host "Setting no space in selection"
$selection.Style = "No Spacing"
Write-Host "Setting font size for selection"
$selection.font.size = 14
Write-Host "Making selection bold"
$selection.font.bold = 1
Write-Host "Making paragraph in selection"
$selection.typeText("Secret Document: Automation Project")
$selection.TypeParagraph()
$selection.font.size = 11
$selection.typeText("Date: $date")
Write-Host "Saving word file "
$doc.saveas([ref] $filePath, [ref]$saveFormat::wdFormatDocument)
$word.Quit
Remove-Variable word 
#Write-Host "Fetching Msword information "
#cscript "C:\Program Files\Microsoft Office\Office14\ospp.vbs" /dstatus
Write-Host "Checking if file is created "
if (!(Test-Path $filePath))
{
Write-Host "File has not been created "
throw "The file doesn't exist"
}
Write-Host "File is created .Script end here !!!"
$null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$word)
[gc]::Collect()
[gc]::WaitForPendingFinalizers()

######################################################

Custom assert in powershell

# This script contains assert function


# This function AssretInt accept two integer type parameter
# It throws exception if both number are not equal

function AssretInt ([int]$a ,[int]$b){
Write-Host 'Inside assert integer function'
if($a -eq $b)
{
Write-Host 'AssertInt is passing..........!!!'
}
else
{
throw 'AssertInt is failing ..........!!!'
}
}
#Below is the example to call the function
#AssretInt '4' '5'
AssretInt '4' '5'

# This function AssretStr accept two string type parameter
# It throws exception if both strings are not equal

function AssertStr([string] $expected ,[string] $actual){
Write-Host 'Inside assert string function'
$c=$expected.CompareTo($actual)

if($c -eq 0)
{
Write-Host 'AssertString is passing ..........!!!'
}
else
{ 
throw 'AssertString is failing ..........!!!'
}
}
#Below is the example to call the function
#AssertStr 'deepak' 'deepak1'

# This function Assretbool accept only one boolean type parameter
# It throws exception if boolean type parameter is false

function Assertbool([bool] $a){
Write-Host 'Inside assert boolean function'

if($a)
{
Write-Host 'AssertBoolean is passing ..........!!!' 
}
else
{ 
throw 'AssertBoolean is failing ..........!!!'
}
}
#Below is the example to call the function
#Assertbool 0

# This function Assertnotnull accept only one string type parameter
# It throws exception if parameter is null or contain white space

function Assertnotnull([string] $message){
Write-Host 'Inside assert not null function'
IF ([string]::IsNullOrWhitespace($message) )
{
throw 'Assert not null is failing ..........!!!' 
}
else
{ 
Write-Host  'Assert not null is passing ..........!!!'
}
}
#$d=''
#Assertnotnull $d



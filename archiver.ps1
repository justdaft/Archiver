
# Created on:   18/09/2012
# Created by:   billyw01
# Filename:     archiver
#
# edit: 10102012 fixed directory name not being used in excel
# edit: 10102012 added copy-spreadsheet function

cls

Reset-Enviroment

$Global:VerbosePreference = 'continue'
$Global:DebugPreference = 'continue'

#Load Required Assemblies
Add-Type -assemblyName PresentationFramework
Add-Type -assemblyName PresentationCore
Add-Type -assemblyName WindowsBase

# Add all global variables.
$ScriptPath = (Split-Path ((Get-Variable MyInvocation).Value).MyCommand.Path)
$GlobalVariables = $ScriptPath + "\GlobalVariables.ps1"

[Windows.Input.InputEventHandler]{ $Global:Window.UpdateLayout() }

function Copy-SpreadSheet($fileName,$sourceLocation=$pwd.Path,$destiantionLocation,[switch]$backup)
{   
    # example  Copy-SpreadSheet -fileName hello.txt -sourceLocation C:\gatewayTest\LatestArchive\  -destiantionLocation C:\gatewayTest\LatestArchive\new\ -backup

    push-location $sourceLocation
    Write-Debug $sourceLocation

    if($backup){
        Write-Debug "backup old file"
        push-location $destiantionLocation
        Write-Debug $destiantionLocation$fileName
        if(Test-Path "$destiantionLocation$("backup_"+$fileName)"){
            Write-Debug "removing old backup $("$destiantionLocation$("backup_"+$fileName)")"
            Remove-Item $destiantionLocation$("backup_"+$fileName)
        }
        Rename-Item $destiantionLocation$fileName  $destiantionLocation$("backup_"+$fileName) -Force
        Pop-Location
    }
    copy-item $fileName $destiantionLocation$fileName
    Pop-Location 
}

Function Calculate-MD5 
{
    Param([Parameter(Mandatory = $True,
        ValueFromPipeLine = $True,
        Position = 0)]
        [String]$Filename
    )
    # Must include full path to file
    $MD5 = New-Object System.Security.Cryptography.MD5CryptoServiceProvider
    If([System.IO.File]::Exists($FileName)) {
        write-debug "Hashing $filename"
        $FileStream = New-Object System.IO.FileStream($FileName,`
        [System.IO.FileMode]::Open,[System.IO.FileAccess]::Read,`
        [System.IO.FileShare]::ReadWrite)
        [byte[]]$ByteSum = $MD5.ComputeHash($FileStream)
        $Hash = ([System.Bitconverter]::ToString($ByteSum)).Replace("-","")
        $FileStream.Close()
    }
    Else {
        $HASH = "ERROR: $FileName Not Found"
    }
    write-debug "Hash for $filename is $($HASH) "
    return $hash
}

function New-archive 
{
    param (
        [string]$directory,        
        [string]$logName,
        [string]$logSize,
        [string]$logLastWriteTime,
        [string]$archiver,
        [string]$date,
        [string]$md5Hash
    )

    New-Object PSObject -Property @{
        Directory = $directory
        Name = $logName
        Size = $logSize
        LastWrite = $logLastWriteTime
        Archiver = $archiver
        Date = $Date
        MD5Hash = $md5Hash
    }
} #end function

function Copy-Archive 
{
	[CmdletBinding()]
	param(
		[Parameter(Position=0, Mandatory=$true)]
		[ValidateNotNullOrEmpty()]
		[System.String]
		$source,
		[Parameter(Position=1, Mandatory=$true)]
		[ValidateNotNullOrEmpty()]
		[System.String]
		$destination,
		[Parameter(Position=2, Mandatory=$true)]
		[ValidateNotNullOrEmpty()]
		[System.String]
		$name
	)
	BEGIN {

        Set-Location $source
        new-item -ItemType D -Path $destination$Name
        $FileSource = gci $source

    }

    PROCESS {
    foreach ($file in $FileSource)
        {
            Copy-Item $File  -Destination "$destination$Name"
            
        }
	}
} #end function

function Get-ArchiveName
{
	[CmdletBinding()]
	param(
		[Parameter(Position=0, Mandatory=$true)]
		[ValidateNotNullOrEmpty()]
		[System.String]
		$filePath
	)
	PROCESS {
			set-location $filePath
			[string]$fullName= Get-childitem *.log | Where-Object {$_.Name -match "ArchiveAudit\d{1,8}"}
			$ArchiveName = $fullName.Substring($fullName.Length – 12, 8)
			return $ArchiveName
}
} #end function

function Update-Stats ()
{
    $listboxOnSiteStats.Items.Clear()
    $listboxOffSiteStats.Items.Clear()
    $onSiteFolderCount = (gci $onSiteStorage  | ? { $_.psiscontainer } ).count
    $onSiteFileCount = (gci $onSiteStorage  -Recurse | ? { !$_.psiscontainer } ).count
    $offSiteFolderCount = (gci $offSiteStorage  | ? { $_.psiscontainer} ).count
    $offSiteFileCount = (gci $offSiteStorage  -Recurse | ? { !$_.psiscontainer }).count
    $listboxOnSiteStats.Items.Add("number of folders: $onSiteFolderCount")
    $listboxOnSiteStats.Items.Add("number of files: $onSiteFileCount")
    $listboxOffSiteStats.Items.Add("number of folders: $offSiteFolderCount")
    $listboxOffSiteStats.Items.Add("number of files: $offSiteFileCount") 
} #end function

function Upate-List0ffSite ()
{
    [array]$listOffSite = gci $offSiteStorage | select -ExpandProperty name
    $listboxOffSite.ItemsSource = $listOffSite
} #end function

function Update-List0nSite ()
{
    [array]$listOnSite = gci $onSiteStorage | select -ExpandProperty name
    $listboxOnSite.ItemsSource = $listOnSite
} #end function

function Update-ListLocal ()
{
    [array]$list = gci $latestLogs | select -ExpandProperty name
    $listboxLocal.ItemsSource = $list
} #end function

function Update-List ($listBox,$dirPath)
{
    [array]$list = gci $dirPath | select -ExpandProperty name
    $listBox.ItemsSource = $list
} #end function


#region Setup
    Set-Location $ScriptPath
    $GLOBAL:list = ""
    Write-Debug "GLOBAL Variables, start"
    . $GlobalVariables
    Write-Debug "`t GLOBAL:user $GLOBAL:user "
    Write-Debug "`t GLOBAL:onSiteStorage $GLOBAL:onSiteStorage"
    Write-Debug "`t GLOBAL:offSiteStorage $GLOBAL:offSiteStorage"
    Write-Debug "`t GLOBAL:latestLogs $GLOBAL:latestLogs"
    Write-Debug "`t GLOBAL:smtpServer $GLOBAL:smtpServer "
    Write-Debug "`t GLOBAL:emailFrom $GLOBAL:emailFrom"
    Write-Debug "`t GLOBAL:emailSubject $GLOBAL:emailSubject"
    Write-Debug "`t GLOBAL:sendEmail $GLOBAL:sendEmail"
    Write-Debug "`t GLOBAL:SendAttachment $GLOBAL:SendAttachment"
    Write-Debug "`t GLOBAL:excelFile $GLOBAL:excelFile"
    Write-Debug "GLOBAL Variables, end `r`n"
#endregion

#region xaml
    [XML]$xaml=gc ".\MainWindow.xaml"
#endregion

#region FindName
    $GLOBAL:reader=(New-Object System.Xml.XmlNodeReader $XAML)

    $GLOBAL:Window=[Windows.Markup.XamlReader]::Load( $reader )

    $GLOBAL:listboxLocal = $GLOBAL:Window.FindName("listboxLocal")
    $GLOBAL:listboxOffsite = $GLOBAL:Window.FindName("listboxOffSite")
    $GLOBAL:listboxOffsiteFiles = $GLOBAL:Window.FindName("listboxOffSiteFiles")
    $GLOBAL:listboxOnSite = $GLOBAL:Window.FindName("listboxOnSite")
    $GLOBAL:listboxOnSiteFiles = $GLOBAL:Window.FindName("listboxOnSiteFiles")
    $GLOBAL:listBoxTasks = $GLOBAL:Window.FindName("listBoxTasks")
    $GLOBAL:listboxOnSiteStats = $GLOBAL:Window.FindName("listboxOnSiteStats")
    $GLOBAL:listboxOffSiteStats = $GLOBAL:Window.FindName("listboxOffSiteStats")

        
    $GLOBAL:lbItem1 = $GLOBAL:Window.FindName("lbItem1")
    $GLOBAL:lbItem2 = $GLOBAL:Window.FindName("lbItem2")
    $GLOBAL:lbItem3 = $GLOBAL:Window.FindName("lbItem3")
    $GLOBAL:lbItem4 = $GLOBAL:Window.FindName("lbItem4")
    $GLOBAL:lbItem5 = $GLOBAL:Window.FindName("lbItem5")
    $GLOBAL:lbItem6 = $GLOBAL:Window.FindName("lbItem6")
    $GLOBAL:lbItem7 = $GLOBAL:Window.FindName("lbItem7")

    $GLOBAL:buttonCreateArchive = $GLOBAL:Window.FindName("buttonCreateArchive")
    $GLOBAL:buttonViewExcel = $GLOBAL:Window.FindName("buttonViewExcel")

    $GLOBAL:labelUser = $GLOBAL:Window.FindName("labelUser")
    $GLOBAL:onsiteArchive = $GLOBAL:Window.FindName("onsiteArchive")
    $GLOBAL:offsiteArchive = $GLOBAL:Window.FindName("offsiteArchive")
    $GLOBAL:textblockStatus = $GLOBAL:Window.FindName("textblockStatus")

#endregion

#Window Load Events

$GLOBAL:Window.WindowStartupLocation = "CenterScreen"

$GLOBAL:Window.Add_Loaded({
    set-location $latestLogs
    $labelUser.content = $user

    Update-ListLocal
    Upate-List0ffSite
    Update-List0nSite
    Update-Stats
    $buttonViewExcel.IsEnabled = $false
    ##Configure a timer to refresh window##
    #Create Timer object
    Write-Verbose "Creating timer object"
    $Global:timer = new-object System.Windows.Threading.DispatcherTimer 
    #Fire off every 2 seconds
    Write-Verbose "Adding 2 second interval to timer object"
    $timer.Interval = [TimeSpan]"0:0:2.00"
    #Add event per tick
    Write-Verbose "Adding Tick Event to timer object"
    $timer.Add_Tick({
        [Windows.Input.InputEventHandler]{ $Global:Window.UpdateLayout() }
        $Global:Window.Dispatcher.Invoke( "Render", [Windows.Input.InputEventHandler]{ $Global:Window.UpdateLayout() }, $null, $null)
        Write-Debug "update layout"
    })
    #Start timer
    Write-Verbose "Starting Timer"
    $timer.Start()
    If (-NOT $timer.IsEnabled) {
        $Window.Close()
    }
})   

$buttonCreateArchive.add_click({
    Set-Location $latestLogs
    $listboxOnSite.ItemsSource = $null
    $textblockStatus.Text = "Running"
    $fileList = GCI $latestLogs


    Write-Debug "Getting Archive Folder name "
    $textblockStatus.Text = "Getting Archive Folder name"    		
	$folderName = Get-ArchiveName -filePath $latestLogs

	Write-Debug "New Archive Folder $folderName"
    $textblockStatus.Text = "New Archive Folder $folderName"  
    $textblockStatus.Text = "New Archive Folder Path $folderName"
	$newArchiveFolderPath =$latestLogs+$folderName   

	Write-Debug "Create New Archive" 
	$lbItem1.content = "1: CREATE NEW ONSITE FOLDER           DONE"
	$lbItem1.Background="#FF00FF00"
    $Global:Window.Dispatcher.Invoke( "Render", [Windows.Input.InputEventHandler]{ $Global:Window.UpdateLayout() }, $null, $null)
    Start-Sleep 1
	
	
    Write-Debug "New Archive Created $folderName" 
    $archive = @()
    $all =  gci   |  select name,length,LastWriteTime,@{Name="MD5Hash";Expression={ $(Calculate-MD5 $_.fullname)}}
    $all |%{ $archive += New-archive -directory $folderName -logname $_.name -logSize $_.length -logLastWriteTime $_.lastwritetime -archiver $user -date $Date -md5Hash $_.MD5Hash}   
    

    #$FileSource = gci $latestLogs 
	$lbItem2.content = "2: COPY TO ONSITE STORAGE               DONE"
	$lbItem2.Background="#FF00FF00"
    $Global:Window.Dispatcher.Invoke( "Render", [Windows.Input.InputEventHandler]{ $Global:Window.UpdateLayout() }, $null, $null)
    Start-Sleep 1

    Write-Debug "copy archive to onsite storage " 
    Copy-Archive -source $latestLogs -destination $onSiteStorage -name $folderName
    Update-List0nSite

    $onsiteArchive.content = "Added $folderName"
       if (test-path $onSiteStorage$folderName){
		$lbItem3.content = "3: CREATE NEW OFFSITE FOLDER          DONE"
		$lbItem3.Background="#FF00FF00"
        $Global:Window.Dispatcher.Invoke( "Render", [Windows.Input.InputEventHandler]{ $Global:Window.UpdateLayout() }, $null, $null)
        Start-Sleep 1
    }

    Write-Debug "copy archive to offsite storage "  
    Copy-Archive -source $latestLogs -destination $oFFSiteStorage -name $folderName
    Upate-List0ffSite

    $offsiteArchive.content = "Added $folderName"
       if (test-path $offSiteStorage$folderName){
	$lbItem4.content = "4: COPY TO OFFSITE STORAGE              DONE"
		$lbItem4.Background="#FF00FF00"
        $Global:Window.Dispatcher.Invoke( "Render", [Windows.Input.InputEventHandler]{ $Global:Window.UpdateLayout() }, $null, $null)
        Start-Sleep 1
    }

    Write-Debug "Cleaning up $latestLogs"
	Set-Location $latestLogs
    Remove-Item *.log 
    Update-ListLocal
    if (!$(test-path $latestLogs$folderName)){
	    $lbItem5.content = "5: REMOVE ORIGINAL LOGS                  DONE"
		$lbItem5.Background="#FF00FF00"
        $Global:Window.Dispatcher.Invoke( "Render", [Windows.Input.InputEventHandler]{ $Global:Window.UpdateLayout() }, $null, $null)
        Start-Sleep 1
    }

    Write-Debug "Creating Email"

    $GLOBAL:arclist = ""

    foreach ($file in $fileList) {
		    #$arclist += "`n"
		    $arclist += "`t$file`r`n"
			}

    $emailBody = @"
    $emailHeader

    Archive Name: $folderName

    Files: $arclist

    $emailFooter
"@

 
	if( $GLOBAL:sendEmail){
		Write-Debug "Sending Email"

    for ( [int]$attempt = 1; $attempt -le 3; $attempt++ )
        { 
            [bool]$success = $false;

            try
                {
                    send-mailmessage -to $emailToList -from $emailFrom -body $emailBody -subject $emailSubject -SmtpServer $smtpServer
                    $success = $true;
                    Write-Debug "Send Email Attempt $attempt succeeded.";
                }
            catch [System.Exception]
                {
                    Write-Debug "$("-" * 80)`r`nSend Email Attempt $attempt failed:`r`n"
                    Write-Debug $Error[0]
                    Write-Debug $("-" * 80); 
                    Write-Debug "`r`n"
                }

              # If the message succeeded, exit the loop
              if($success) { break; }
        }   	
	}
	else
	{
		Write-Debug "Not Sending Email" 
	}
    Write-Debug "Email Sent" 
	$lbItem6.content = "6: SEND EMAIL                                       DONE"
	$lbItem6.Background="#FF00FF00"
    $Global:Window.Dispatcher.Invoke( "Render", [Windows.Input.InputEventHandler]{ $Global:Window.UpdateLayout() }, $null, $null)
    Start-Sleep 1

& {    		
    Write-Debug "Updateing excel"
    $GLOBAL:xl = New-Object -ComObject Excel.Application 
    $GLOBAL:wb = $xl.Workbooks.Open($excelFile)
    Write-Debug "objects $GLOBAL:xl and $GLOBAL:wb, created"
        
    #$xl.Visible = $true

    $ws = $wb.Worksheets.Item(1)

    $currentSheet = $wb.Sheets  | where {$_.name -eq "Archives"}

    $currentSheet.activate()

    $mainRng = $currentSheet.UsedRange.Cells 
    $RowCount = $mainRng.Rows.Count  
    $R = $RowCount
    $inc = $R + 1

    foreach ($item in $archive)
        {
            $currentSheet.cells.Item($inc,1) = $item.Directory
            $currentSheet.cells.Item($inc,2) = $item.LastWrite 
            $currentSheet.cells.Item($inc,3) = $item.Name 
            $currentSheet.cells.Item($inc,4) = $item.Size 
            $currentSheet.cells.Item($inc,5) = $item.Archiver 
            $currentSheet.cells.Item($inc,6) = $item.Date
            $currentSheet.cells.Item($inc,7) = $item.MD5Hash
            $inc ++
        }
    $wb.Save() 
    $xl.quit()

    $onsiteReport = $onSiteStorage+"ArchiveReport\"
    Write-Debug "copying spreadsheet to $onsiteReport"
    Copy-SpreadSheet -fileName ArchiveDetails.xlsx -sourceLocation C:\gatewayTest\ArchiveReport\  -destiantionLocation $onsiteReport -backup

    
    $offsiteReport = $offSiteStorage+"ArchiveReport\"
    Write-Debug "copying spreadsheet to $offsiteReport"
    Copy-SpreadSheet -fileName ArchiveDetails.xlsx -sourceLocation C:\gatewayTest\ArchiveReport\  -destiantionLocation $offsiteReport -backup

	#if any processes left, kill them
	
	$xl | ForEach {[void][Runtime.Interopservices.Marshal]::ReleaseComObject($_)}
   # if (ps excel) { kill -name excel}
	}
	
	$lbItem7.content = "7: UPDATE EXCEL REPORT                     DONE"
	$lbItem7.Background="#FF00FF00"

    Update-Stats

    $buttonCreateArchive.IsEnabled = $false
    $buttonViewExcel.IsEnabled = $True
    $updateWindow
})

$buttonViewExcel.add_click({
			ii "C:\gatewayTest\ArchiveReport\ArchiveDetails.xlsx"
	})

$GLOBAL:listboxOffSite.add_SelectionChanged({
        write-debug $GLOBAL:listboxOffsite.SelectedValue
        $searchPath = $offSiteStorage+$GLOBAL:listboxOffsite.SelectedValue
        Update-List -listBox $GLOBAL:listboxOffsiteFiles -dirPath $searchPath
        $Global:Window.Dispatcher.Invoke( "Render", [Windows.Input.InputEventHandler]{ $Global:Window.UpdateLayout() }, $null, $null)
    })

$GLOBAL:listboxOnSite.add_SelectionChanged({
        write-debug $GLOBAL:listboxOnSite.SelectedValue
        $searchPath = $onSiteStorage+$GLOBAL:listboxOnSite.SelectedValue
        Update-List -listBox $GLOBAL:listboxOnSiteFiles -dirPath $searchPath
        $Global:Window.Dispatcher.Invoke( "Render", [Windows.Input.InputEventHandler]{ $Global:Window.UpdateLayout() }, $null, $null)
    })

#Window Events
$Global:Window.Add_Closed({

    $timer.Stop() 
})

$GLOBAL:Window.ShowDialog() | out-null

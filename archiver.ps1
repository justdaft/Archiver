
# Created on:   18/09/2012 13:46
# Created by:   billyw01
# Filename:     archiver

cls
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


Function Calculate-MD5 {
    Param([Parameter(Mandatory = $True,
        ValueFromPipeLine = $True,
        Position = 0)]
        [String]$Filename
    )

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
}

function reset($source='C:\gatewayTest\logs\', $destination='c:\gatewayTest\LatestArchive\'){
    push-location $source
    
    $random = Get-Random  -Maximum 99999999 -Minimum 11111111
    gci | Where-Object {$_.Name -match "ArchiveAudit\d{1,8}"} | rename-item -NewName .\ArchiveAudit$random.log
    gci | Where-Object {$_.Name -match "ArchiveIsdSyncAudit\d{1,8}"} | rename-item -NewName .\ArchiveIsdSyncAudit$random.log
    gci $source  *.log | copy-Item  -Destination $destination
    
    Pop-Location
} #end function

function New-archive {
param (
[string]$directoryName,
[string]$logName,
[string]$logSize,
[string]$logLastWriteTime,
[string]$archiver,
[string]$date
)

New-Object PSObject -Property @{
    Directory = $directoryName
    Name = $logName
    Size = $logSize
    LastWrite = $logLastWriteTime
    Archiver = $archiver
    Date = $Date
}
} #end function

function Copy-Archive {
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

function Get-ArchiveName {
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


function updateStats (){
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



function upDatelist0ffSite ()
{
    [array]$listOffSite = gci $offSiteStorage | select -ExpandProperty name
    $listboxOffSite.ItemsSource = $listOffSite
} #end function


function upDatelist0nSite ()
{
    [array]$listOnSite = gci $onSiteStorage | select -ExpandProperty name
    $listboxOnSite.ItemsSource = $listOnSite
} #end function


function upDatelistLocal ()
{
    [array]$list = gci $latestLogs | select -ExpandProperty name
    $listboxLocal.ItemsSource = $list
} #end function

function upDatelist ($listBox,$dirPath)
{
    [array]$list = gci $dirPath | select -ExpandProperty name
    $listBox.ItemsSource = $list
} #end function


#region Setup
reset
$GLOBAL:list = ""
. $GlobalVariables


Write-Debug "GLOBAL:user $GLOBAL:user "
Write-Debug "GLOBAL:onSiteStorage $GLOBAL:onSiteStorage"
Write-Debug "GLOBAL:offSiteStorage $GLOBAL:offSiteStorage"
Write-Debug "GLOBAL:latestLogs $GLOBAL:latestLogs"
Write-Debug "GLOBAL:smtpServer $GLOBAL:smtpServer "
Write-Debug "GLOBAL:emailFrom $GLOBAL:emailFrom"
Write-Debug "GLOBAL:emailSubject $GLOBAL:emailSubject"
Write-Debug "GLOBAL:sendEmail $GLOBAL:sendEmail"
Write-Debug "GLOBAL:SendAttachment $GLOBAL:SendAttachment"
Write-Debug "GLOBAL:excelFile $GLOBAL:excelFile"

#endregion

#region xaml
    [XML]$xaml=gc "C:\Dropbox\Repositories\archiver\archiver\MainWindow.xaml"
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

    upDatelistLocal
    upDatelist0ffSite
    upDatelist0nSite
    updateStats


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
    $all =  gci   |  select name,length,LastWriteTime
    $all |%{ $archive += New-archive -directoryName $folderName -logname $_.name -logSize $_.length -logLastWriteTime $_.lastwritetime -archiver $user -date $Date }    
    
    #$FileSource = gci $latestLogs 
	$lbItem2.content = "2: COPY TO ONSITE STORAGE               DONE"
	$lbItem2.Background="#FF00FF00"
    $Global:Window.Dispatcher.Invoke( "Render", [Windows.Input.InputEventHandler]{ $Global:Window.UpdateLayout() }, $null, $null)
    Start-Sleep 1

    Write-Debug "copy archive to onsite storage " 
    Copy-Archive -source $latestLogs -destination $onSiteStorage -name $folderName
    upDatelist0nSite

    $onsiteArchive.content = "Added $folderName"
       if (test-path $onSiteStorage$folderName){
		$lbItem3.content = "3: CREATE NEW OFFSITE FOLDER          DONE"
		$lbItem3.Background="#FF00FF00"
        $Global:Window.Dispatcher.Invoke( "Render", [Windows.Input.InputEventHandler]{ $Global:Window.UpdateLayout() }, $null, $null)
        Start-Sleep 1
    }

    Write-Debug "copy archive to offsite storage "  
    Copy-Archive -source $latestLogs -destination $oFFSiteStorage -name $folderName
    upDatelist0ffSite

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
    upDatelistLocal
    if (!$(test-path $latestLogs$folderName)){
	    $lbItem5.content = "5: REMOVE ORIGINAL LOGS                  DONE"
		$lbItem5.Background="#FF00FF00"
        $Global:Window.Dispatcher.Invoke( "Render", [Windows.Input.InputEventHandler]{ $Global:Window.UpdateLayout() }, $null, $null)
        Start-Sleep 1
    }

    Write-Debug "Creating Email"
    foreach ($file in $fileList) {
		    $list += "`n"
		    $list += "`t`t$file"
			}


    
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
            $inc ++
        }
    $wb.Save() 
    $xl.quit()


	#if any processes left, kill them
	
	$xl | ForEach {[void][Runtime.Interopservices.Marshal]::ReleaseComObject($_)}
   # if (ps excel) { kill -name excel}
	}
	
	$lbItem7.content = "7: UPDATE EXCEL REPORT                     DONE"
	$lbItem7.Background="#FF00FF00"


updateStats


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
        updatelist -listBox $GLOBAL:listboxOffsiteFiles -dirPath $searchPath
        $Global:Window.Dispatcher.Invoke( "Render", [Windows.Input.InputEventHandler]{ $Global:Window.UpdateLayout() }, $null, $null)
    })

$GLOBAL:listboxOnSite.add_SelectionChanged({
        write-debug $GLOBAL:listboxOnSite.SelectedValue
        $searchPath = $onSiteStorage+$GLOBAL:listboxOnSite.SelectedValue
        updatelist -listBox $GLOBAL:listboxOnSiteFiles -dirPath $searchPath
        $Global:Window.Dispatcher.Invoke( "Render", [Windows.Input.InputEventHandler]{ $Global:Window.UpdateLayout() }, $null, $null)
    })

#Window Events
$Global:Window.Add_Closed({


    $timer.Stop() 
})

$GLOBAL:Window.ShowDialog() | out-null

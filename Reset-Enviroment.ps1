function Reset-Enviroment($source='C:\gatewayTest\logs\', $destination='c:\gatewayTest\LatestArchive\'){
    push-location $source
    
    $random = Get-Random  -Maximum 99999999 -Minimum 11111111
    gci | Where-Object {$_.Name -match "ArchiveAudit\d{1,8}"} | rename-item -NewName .\ArchiveAudit$random.log
    gci | Where-Object {$_.Name -match "ArchiveSyncAudit\d{1,8}"} | rename-item -NewName .\ArchiveSyncAudit$random.log
    gci $source  *.log | copy-Item  -Destination $destination


    $txt = @()
    foreach ($item in GCI)
    {
       $txt += get-date
       $random1 = Get-Random  -Maximum 99 -Minimum 1
       $text = 1..$random1 | % { $txt += $(Get-LoremIpsum) }
       set-content $item -Value $txt
    }
    
    Pop-Location
} #end function
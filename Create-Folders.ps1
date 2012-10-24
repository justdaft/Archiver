function Create-folders ($file=".\FoldersLIst.txt", $root=".\")
{
    $collection = Get-Content $file
    Write-Debug "using `$collection $collection"
    foreach ($item in $collection)
    {
       write-debug "creating `$root$item $root$item"
       new-item -ItemType directory $root$item 
    }
    
       
}

Create-folders -file .\FoldersLIst.txt -root c:\test\

#Select folder to convert From
$RootFolder = "C:\temp\Convert\"
#If documents should duplicate original in case it breaks
$ArchiveDoc = $True
#Select folder to duplicate files to in case it breaks
$ArchiveFolder = "C:\temp\Archive\"

#No Changes below this point
$DocFixedFormat = [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatXMLDocument
$DocFixedFormatMacro = [Microsoft.Office.Interop.Word.WdSaveFormat]::wdFormatFlatXMLMacroEnabled
$FileTypeOld = "*doc"
$FileTypeNew = ".docx"
$FileTypeNewMacro = ".docm"
$RootFolderLength = $RootFolder.Length

$excel = New-Object -ComObject word.application
$excel.visible = $false

Get-ChildItem -Path $RootFolder -Include $FileTypeOld -Recurse | ForEach-Object {
    $doc = $_
    $FileName = $doc.Name
    $Path = $doc.FullName
    
    if($ArchiveDoc) {
        #Setup Archive information
        $ArchiveSubFolder = $ArchiveFolder+$Path.Substring($RootFolderLength, $Path.Length-$RootFolderLength-$FileName.Length)
        $ArchivePath = $ArchiveSubFolder+$FileName
        #Archives original
        if(-not(Test-Path -PathType Container $ArchiveSubFolder)) {
            New-Item -item Directory -Force -Path $ArchiveSubFolder | Out-Null
        }
        if(-not(Test-Path -PathType Leaf $ArchivePath)){
            Copy-Item -Path $Path -Destination $ArchivePath
        } else {
            Write-Host "File: "$Path "Exists in archive already" 
        }
    }

    #Convert file
    $workbook = $word.Workbooks.Open($Path)
    if($workbook.HasVBProject) {
        $PathNew = $Path.Substring(0, $Path.Length-$FileTypeOld.Length)+$FileTypeNewMacro
        $workbook.SaveAs($PathNew, $DocFixedFormatMacro)
    } else {
        $PathNew = $Path.Substring(0, $Path.Length-$FileTypeOld.Length)+$FileTypeNew
        $workbook.SaveAs($PathNew, $DocFixedFormat)
    }
    $workbook.Close()

    #Remove old file
    Remove-Item $Path
    
    #Clear Temp
    $workbook = $null
}

#Clear Temp
$word.Quit()
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($word)
$word = $null

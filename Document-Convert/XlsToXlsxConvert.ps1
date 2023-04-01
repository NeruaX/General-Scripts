#Select folder to convert From
$RootFolder = "C:\temp\Convert\"
#If documents should duplicate original in case it breaks
$ArchiveXls = $True
#Select folder to duplicate files to in case it breaks
$ArchiveFolder = "C:\temp\Archive\"

#No Changes below this point
$xlFixedFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbook
$xlFixedFormatMacro = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbookMacroEnabled
$FileTypeOld = "*xls"
$FileTypeNew = ".xlsx"
$FileTypeNewMacro = ".xlsm"
$RootFolderLength = $RootFolder.Length

$excel = New-Object -ComObject excel.application
$excel.visible = $false

Get-ChildItem -Path $RootFolder -Include $FileTypeOld -Recurse | ForEach-Object {
    $xls = $_
    $FileName = $xls.Name
    $Path = $xls.FullName
    
    if($ArchiveXls) {
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
    $workbook = $excel.Workbooks.Open($Path)
    if($workbook.HasVBProject) {
        $PathNew = $Path.Substring(0, $Path.Length-$FileTypeOld.Length)+$FileTypeNewMacro
        $workbook.SaveAs($PathNew, $xlFixedFormatMacro)
    } else {
        $PathNew = $Path.Substring(0, $Path.Length-$FileTypeOld.Length)+$FileTypeNew
        $workbook.SaveAs($PathNew, $xlFixedFormat)
    }
    $workbook.Close()

    #Remove old file
    Remove-Item $Path
    
    #Clear Temp
    $workbook = $null
}

#Clear Temp
$excel.Quit()
[void][System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
$excel = $null

<#
    .SYNOPSIS
        Converts all XLS files in a provided directory to XLSX.
    .DESCRIPTION
        Just drop a folder containing (PubStat) XLS files on the bat file
        and it will convert them all to XLSX, while keeping a
        copy in an xls folder it creates. Yes, it works recursively.
        
        Largely inspired by: https://gist.github.com/gabceb/954418
    .PARAMETER folderPath
        A directory that contains XLS files
#>
[CmdletBinding()]
param (
    [Parameter(ValueFromRemainingArguments=$true)]
    $folderPath
)
Add-Type -AssemblyName Microsoft.Office.Interop.Excel

$xlFixedFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlOpenXMLWorkbook
$excel = New-Object -ComObject excel.application
$excel.visible = $false
$fileFormat = "xls"

try {
    # Get XLS files in the folder hierarchy, but exclude _xls 
    # directories (which contain backups of previously converted XLS files)
    # https://stackoverflow.com/questions/15294836/how-can-i-exclude-multiple-folders-using-get-childitem-exclude
    $xlsFiles = Get-ChildItem -Path $folderpath -Include "*.$fileFormat" -recurse |
        Where-Object { $_.FullName -inotmatch "_$fileFormat" }
    
    $fileCount = ($xlsFiles | Measure-Object).Count
    Write-Host  "$fileCount $fileFormat files found." 
    $xlsFiles | ForEach-Object `
    {
        $path = ($_.FullName).substring(0, ($_.FullName).lastindexOf("."))
        "Converting $path ..."
        $workbook = $excel.workbooks.open($_.FullName)
        $path += ".xlsx"
        If (Test-Path $path) 
        {
            Remove-Item $Path
        }
        $workbook.SaveAs($path, $xlFixedFormat)
        $workbook.Close()

        # Move the original XLS file into the xls folder
        $xlsFolder = $path.substring(0, $path.lastIndexOf("\")) + "\_$fileFormat"
	    If(-not (test-path $xlsFolder))
        {
            New-Item $xlsFolder -type directory | Out-Null
        }
        Move-Item $_.FullName $xlsFolder
    }
    If ($count -gt 0) 
    {
        Write-Host "Any original files have been saved in _xls folders." 
    }
}
# If you do a Ctrl+C this will also be called, so no locked files that cannot be deleted!
finally {
    $excel.Quit()
    $excel = $null
    [gc]::collect()
    [gc]::WaitForPendingFinalizers() 
    Write-Host "Process finished."
}

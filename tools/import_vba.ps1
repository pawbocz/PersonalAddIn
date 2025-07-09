Param(
    [string]$xlam = "..\bin\PersonalAddIn.xlam",
    [string]$src  = "..\src"
)

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$wb = $excel.Workbooks.Open((Resolve-Path $xlam))

# wyczyść projekt
foreach ($comp in @($wb.VBProject.VBComponents)) {
    $wb.VBProject.VBComponents.Remove($comp)
}

# importuj pliki
Get-ChildItem $src -Filter *.bas,*.cls,*.frm | ForEach-Object {
    $wb.VBProject.VBComponents.Import($_.FullName)
}

$wb.Save()
$wb.Close()
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

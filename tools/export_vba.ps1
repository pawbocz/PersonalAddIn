Param(
    [string]$xlam = "..\bin\PersonalAddIn.xlam",
    [string]$out  = "..\src"
)

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

$addinName = (Split-Path $xlam -Leaf)
$addinObj  = $excel.AddIns | Where-Object { $_.Name -eq $addinName }
if ($addinObj) { $addinObj.Installed = $false }

$wb = $excel.Workbooks.Open((Resolve-Path $xlam))

if (-not (Test-Path $out)) { New-Item -ItemType Directory -Path $out | Out-Null }
Remove-Item "$out\*" -Force -ErrorAction SilentlyContinue

foreach ($comp in $wb.VBProject.VBComponents) {

    switch ($comp.Type) {
        1 { $ext = "bas" }   # vbext_ct_StdModule
        2 { $ext = "cls" }   # vbext_ct_ClassModule
        3 { $ext = "frm" }   # vbext_ct_MSForm   (.frm + .frx)
        Default { $ext = "bas" }
    }

    $path = Join-Path $out ("{0}.{1}" -f $comp.Name, $ext)
    $comp.Export($path)

}

$wb.Close($false)
$excel.Quit()

[System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb)    | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()

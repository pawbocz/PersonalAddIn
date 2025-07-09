Param(
    # pełne ścieżki zbudowane na podstawie PSScriptRoot (folderu, w którym jest ten plik)
    [string]$xlam = (Join-Path $PSScriptRoot '..\bin\PersonalAddIn.xlam'),
    [string]$out  = (Join-Path $PSScriptRoot '..\src')
)

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# upewnij się, że plik istnieje
if (-not (Test-Path $xlam)) {
    Write-Error "Nie znaleziono pliku add-in: $xlam"
    exit 1
}

# utwórz src, jeśli go nie ma
if (-not (Test-Path $out)) { New-Item -ItemType Directory -Path $out | Out-Null }

# (1) odłącz dodatek, jeśli był załadowany w tej instancji Excela
$addinName = [IO.Path]::GetFileName($xlam)
$addinObj  = $excel.AddIns | Where-Object { $_.Name -eq $addinName }
if ($addinObj) { $addinObj.Installed = $false }

# (2) otwórz plik i eksportuj
$wb = $excel.Workbooks.Open($xlam)

Remove-Item "$out\*" -Force -ErrorAction SilentlyContinue

foreach ($comp in $wb.VBProject.VBComponents) {
    switch ($comp.Type) {
        1 { $ext = 'bas' }   # standard module
        2 { $ext = 'cls' }   # class module
        3 { $ext = 'frm' }   # userform
        default { $ext = 'bas' }
    }
    $target = Join-Path $out ("$($comp.Name).$ext")
    $comp.Export($target)
}

$wb.Close($false)
$excel.Quit()
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb)    | Out-Null
[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
[GC]::Collect(); [GC]::WaitForPendingFinalizers()

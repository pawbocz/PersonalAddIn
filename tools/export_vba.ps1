Param(
    # Lista plików do eksportu. Każdy ma osobny katalog src, więc nic się nie miesza.
    [Parameter(Mandatory = $false)]
    [array]$targets = @(
        @{ xlam = (Join-Path $PSScriptRoot '..\bin\PersonalAddIn.xlam'); out = (Join-Path $PSScriptRoot '..\src\PersonalAddIn') }
        # <-- DODAJ TU NOWY PLIK, np.:
        # @{ xlam = (Join-Path $PSScriptRoot '..\bin\MegaLV.xlsm');     out = (Join-Path $PSScriptRoot '..\src\QF_CopyToLV') }
    )
)

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    foreach ($t in $targets) {
        $xlam = $t.xlam
        $out  = $t.out

        if (-not (Test-Path $xlam)) { Write-Error "Nie znaleziono pliku: $xlam"; throw "Brak pliku" }
        if (-not (Test-Path $out)) { New-Item -ItemType Directory -Force -Path $out | Out-Null }

        # odłącz dodatek, jeśli załadowany w tej instancji
        $addinName = [IO.Path]::GetFileName($xlam)
        $addinObj  = $excel.AddIns | Where-Object { $_.Name -eq $addinName }
        if ($addinObj) { $addinObj.Installed = $false }

        $wb = $excel.Workbooks.Open($xlam)

        # wyczyść katalog docelowy TYLKO dla tego projektu
        Get-ChildItem -Path $out -File -ErrorAction SilentlyContinue | Remove-Item -Force -ErrorAction SilentlyContinue

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
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
    }
}
finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

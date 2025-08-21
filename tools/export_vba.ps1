Param(
    # Opcjonalnie: własna lista targetów przekazana z zewnątrz.
    # Jeśli pusta, skrypt sam zbuduje listę dla PersonalAddIn.xlam i QF_CopyToLV.xlam.
    [array]$targets
)

# --- Ustal bazową ścieżkę na podstawie położenia skryptu ---
$Base = Split-Path -Parent $PSCommandPath

# --- Domyślne targety, jeśli nie podano w Param ---
if (-not $targets -or $targets.Count -eq 0) {
    $targets = @(
        @{ xlam = (Join-Path $Base '..\bin\PersonalAddIn.xlam'); out = (Join-Path $Base '..\src\PersonalAddIn') },
        @{ xlam = (Join-Path $Base '..\bin\QF_CopyToLV.xlam');   out = (Join-Path $Base '..\src\QF_CopyToLV') }
    )
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    foreach ($t in $targets) {
        $xlam = $t.xlam
        $out  = $t.out

        if (-not (Test-Path $xlam)) {
            Write-Error "Nie znaleziono pliku: $xlam"
            throw "Brak pliku"
        }
        if (-not (Test-Path $out)) {
            New-Item -ItemType Directory -Force -Path $out | Out-Null
        }

        # Odłącz dodatek, jeśli załadowany w tej instancji
        $addinName = [IO.Path]::GetFileName($xlam)
        $addinObj  = $excel.AddIns | Where-Object { $_.Name -eq $addinName }
        if ($addinObj) { $addinObj.Installed = $false }

        $wb = $excel.Workbooks.Open($xlam)

        # Wyczyść katalog docelowy dla tego projektu
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
        [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
    }
}
finally {
    $excel.Quit()
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

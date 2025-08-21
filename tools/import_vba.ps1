Param(
    [array]$targets = @(
        @{ xlam = (Join-Path $PSScriptRoot '..\bin\PersonalAddIn.xlam'); src = (Join-Path $PSScriptRoot '..\src\PersonalAddIn') }
        # @{ xlam = (Join-Path $PSScriptRoot '..\bin\MegaLV.xlsm');     src = (Join-Path $PSScriptRoot '..\src\QF_CopyToLV') }
    )
)

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

try {
    foreach ($t in $targets) {
        $xlam = $t.xlam
        $src  = $t.src

        if (-not (Test-Path $xlam)) { Write-Error "Nie znaleziono pliku: $xlam"; throw "Brak pliku" }
        if (-not (Test-Path $src))  { Write-Error "Brak katalogu źródeł: $src"; throw "Brak src" }

        $wb = $excel.Workbooks.Open($xlam)

        # usuń wszystko oprócz modułów dokumentowych (ThisWorkbook/Worksheets)
        foreach ($comp in @($wb.VBProject.VBComponents)) {
            # Type: 1=Std, 2=Class, 3=UserForm, 100=Document
            if ($comp.Type -ne 100) {
                $wb.VBProject.VBComponents.Remove($comp)
            } else {
                # dla dokumentowych można wyczyścić kod (opcjonalnie):
                # $comp.CodeModule.DeleteLines(1, $comp.CodeModule.CountOfLines)
            }
        }

        # import
        Get-ChildItem $src -File | Where-Object { $_.Extension -in '.bas','.cls','.frm' } | ForEach-Object {
            $wb.VBProject.VBComponents.Import($_.FullName) | Out-Null
        }

        $wb.Save()
        $wb.Close()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
    }
}
finally {
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

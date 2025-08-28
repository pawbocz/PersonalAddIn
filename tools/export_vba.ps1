Param([array]$targets)

$Base = Split-Path -Parent $PSCommandPath

if (-not $targets -or $targets.Count -eq 0) {
    $targets = @(
        @{ xlam = (Join-Path $Base "..\bin\PersonalAddIn.xlam"); out = (Join-Path $Base "..\src\PersonalAddIn") },
        @{ xlam = (Join-Path $Base "..\bin\QF_CopyToLV.xlam");   out = (Join-Path $Base "..\src\QF_CopyToLV") }
    )
}

$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false
$excel.AutomationSecurity = 3  # disable macros on open

try {
    foreach ($t in $targets) {
        $xlam = $t.xlam
        $out  = $t.out

        if (-not (Test-Path $xlam)) { Write-Error ("File not found: " + $xlam); throw "Missing file" }
        if (-not (Test-Path $out))  { New-Item -ItemType Directory -Force -Path $out | Out-Null }

        $full = (Resolve-Path $xlam).Path

        # detach add-in if loaded
        $addinObj = $excel.AddIns | Where-Object { $_.FullName -eq $full }
        if ($addinObj) { $addinObj.Installed = $false }

        # Open(FileName, UpdateLinks, ReadOnly)
        $wb = $excel.Workbooks.Open($full, 0, $true)

        # check VBE access
        try { $null = $wb.VBProject.VBComponents } catch {
            $wb.Close($false)
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
            Write-Error "No access to VBProject. Enable 'Trust access to the VBA project object model' in Excel."
            throw
        }
        if ($wb.VBProject.Protection -ne 0) {
            $wb.Close($false)
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
            Write-Error ("VBA project in " + $xlam + " is locked. Export aborted.")
            throw "VBA locked"
        }

        # clean output folder (only this add-in)
        Get-ChildItem -Path $out -Recurse -Force -ErrorAction SilentlyContinue | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue

        foreach ($comp in $wb.VBProject.VBComponents) {
            switch ($comp.Type) {
                1 { $ext = "bas" }  # standard module
                2 { $ext = "cls" }  # class
                3 { $ext = "frm" }  # userform
                default { $ext = "bas" }
            }
            $fname  = $comp.Name + "." + $ext
            $target = Join-Path -Path $out -ChildPath $fname
            $comp.Export($target)
        }

        $wb.Close($false)
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($wb)
    }
}
finally {
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::ReleaseComObject($excel)
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

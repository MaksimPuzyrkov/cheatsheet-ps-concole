Add-Type @"
    using System;
    using System.Runtime.InteropServices;
    public class Tricks {
        [DllImport("user32.dll")]
        public static extern IntPtr GetForegroundWindow();
    }
"@;

# List all processes if Title exists
Get-Process | Where-Object { $_.MainWindowTitle } | format-table id,name,mainwindowtitle -AutoSize;

# Endless loop that reading available sheets, updates hash table with sheets and runs main loop
while ($true) {
    $SheetFilters = @{};

    # Loading sheet filters
    Get-ChildItem -Path .\sheets\* -Include *.txt | Foreach-Object {
        $SheetFilter = Get-Content -Path $_ -First 1;
        $SheetFilters.Add($SheetFilter, $_.FullName);
    }

    # Main loop will exit 
    foreach($i in 1..1) {
        $ForegroundWindow = [tricks]::GetForegroundWindow();
        $ForegroundProcess = get-process | Where-Object { $_.mainwindowhandle -eq $ForegroundWindow };
        #Clear-Host;
        # Search first matching filter pattern and display if found
        $SheetFound = 0;
        $SheetFilters.Keys | ForEach-Object {
            $ProcessSheetFilter = $_;
            if ($ForegroundProcess.mainwindowtitle -match "$ProcessSheetFilter") {
                Write-Host ("Found Matched Process") -ForegroundColor DarkBlue;
                #$Sheet = Get-Content -Path .\sheets\$($SheetFilters[$_]) -TotalCount 10;
                $Sheet = Get-Content -Path $($SheetFilters[$_]) -TotalCount 100;
                $SheetFound = 1;
            }
        }
        Clear-Host;
        if ($SheetFound -eq 1) {
            $Sheet | ForEach-Object {
                $_;
            }
        } else {
            Write-Output ("Sheet for $($ForegroundProcess.mainwindowtitle) is not found");
        }
        Start-Sleep -s 1;
    }
}
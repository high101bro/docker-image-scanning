<#
    .DESCRIPTION
    Compare two Grype CVE CSV files and output reports.
    Dan Komnick (high101bro)
    11 April 2022

    .SYNOPSIS
    Use the Convert-GrypeJsonToCsv.ps1 to convert the Grype scans to a compiled CSV format. 
    Then compare a new CSV scan results with an older one to see New and Remvoved CVEs.
    You can also compile the results to an Excel format which will color code the severity levels for easy reference.
    If outputted in combiled Excel format, you can also have individuals tabs created for each scan.

    .INPUTS
    CSV files that were output from Grype Scans. The expected format is one that was converted from a json file using a helper script.

    .OUTPUTS
    A CSV file that contains all the New CVEs detected since the ReferenceObject.
    A CSV file that contains all the Removed CSVe no longer present in the DifferenceObject.
    A CSV file that contains all the Removed CSVe no longer present in the DifferenceObject, but just the Critical and High severities.    
    Option to output compiled CVE CSV results to an Excel document which color codes the severity levels.

    .PARAMETER ReferenceObject DifferenceObject ReportPath IndividualTabs CompileToExcel

    .EXAMPLE
    ./Compare-CveCsv.ps1 -ReferenceObject ./OlderScanResults.csv -DifferenceObject ./NewerScanResults.csv    

    .EXAMPLE
    ./Compare-CveCsv.ps1 -ReferenceObject ./OlderScanResults.csv -DifferenceObject ./NewerScanResults.csv -CompileToExcel

    .EXAMPLE
    ./Compare-CveCsv.ps1 -ReferenceObject ./OlderScanResults.csv -DifferenceObject ./NewerScanResults.csv -CompileToExcel -IndividualTabs

    .LINK
    https://github.com/high101bro
#>

param(
    $ReferenceObject,
    $DifferenceObject,
    $ReportPath = (Get-Location | Select-Object -ExpandProperty Path),
    [switch]$IndividualTabs,
    [switch]$CompileToExcel
)

$ReferenceObject  = Import-Csv $ReferenceObject
$DifferenceObject = Import-Csv $DifferenceObject

$ReferenceObjectCVEs = @{}
foreach ( $CVE in $ReferenceObject ) {
    $key = $CVE.Id + ' ' + $CVE.Image
    if ( $key -notin $ReferenceObjectCVEs.keys ) {
        $ReferenceObjectCVEs[$key] += $CVE
    }
}

$DifferenceObjectCVEs = @{}
foreach ( $CVE in $DifferenceObject ) {
    $key = $CVE.Id + ' ' + $CVE.Image
    if ( $key -notin $DifferenceObjectCVEs.keys ) {
        $DifferenceObjectCVEs[$key] += $CVE
    }
}


$NewCVEs = @{}
foreach ( $CVE in $DifferenceObject ) {
    $key = $CVE.Id + ' ' + $CVE.Image
    if ( $key -notin $ReferenceObjectCVEs.keys -and $key -notin $NewCVEs.keys ) {
        $NewCVEs[$key] += $CVE
    }
}


$RemovedCVEs = @{}
foreach ( $CVE in $ReferenceObject ) {
    $key = $CVE.Id + ' ' + $CVE.Image
    if ( $key -notin $DifferenceObject.keys -and $key -notin $RemovedCVEs.keys ) {
        $RemovedCVEs[$key] += $CVE
        #Write-Host $key
    }
}


$NewCVEs.GetEnumerator()     | Select-Object -ExpandProperty Value | Export-Csv .\'New-CVEs-Between-Old-And-New-(All-Severities).csv' -NoTypeInformation
$RemovedCVEs.GetEnumerator() | Select-Object -ExpandProperty Value | Export-Csv .\'Removed-CVEs-Between-Old-And-New.csv' -NoTypeInformation
$NewCVEs.GetEnumerator() | 
    Select-Object -ExpandProperty Value | 
    Where-Object {$_.Severity -eq 'High' -or $_.Severity -eq 'Critical' } | 
    Export-Csv -Path 'New-CVEs-Between-Old-And-New-(Just-High-Critical).csv' -NoTypeInformation -Force


if ($CompileToExcel) {
    [string]$PathToCSVFiles = '.'

    try {
        Set-Location $PathToCSVFile
    }
    catch{}

    $CSVs = Get-ChildItem .\* -Include *.csv
    $y = $CSVs.Count
    
    Write-Host "[!] Detected the following CSV files: ($y)"
    foreach ($CSV in $CSVs) {
        Write-Host "   "$CSV.Name
    }

    #creates file name with date/username
#    $OutputXlsxFileName = "CVE_Container_Scan_Results_" + $(Get-Date -f yyyyMMdd HHmm) + ".xlsx"
    $OutputXlsxFileName = "CVE_Container_Scan_Results_" + "$((Get-Date).ToString('yyyy-MM-dd_HHmm'))" + ".xlsx"
    
    Write-Host "[!] Creating Combine Excel Document:`n    $OutputXlsxFileName"

    $excelapp = New-Object -ComObject Excel.Application
    $excelapp.SheetsInNewWorkbook = $CSVs.Count + 1

    $xlsx = $excelapp.Workbooks.Add()
    $Sheet = 2

    # The first Sheet
    $CompiledWorkSheet = $xlsx.Worksheets.Item(1)
    $CompiledWorkSheet.Name = "Compiled CVEs"
    $CompiledRow = 1

    Write-Host "[!] Combining And Processing CSV Files:"
    foreach ($CSV in $CSVs) {
        Write-Host "    $CSV"
    
        #Adds the Contents of each CSV file to a new sheet        
        $Row = 1
        $Column = 1

        # The first loop is the second Sheet
        $EachWorkSheet     = $xlsx.Worksheets.Item($Sheet)
        $Sheet += 1

        if ($IndividualTabs) {
            $SheetName = $CSV.BaseName -replace '\:','' -replace '\\','' -replace '\/','' -replace '\?','' -replace '\*','' -replace '\[','' -replace '\]',''
            $SheetName = $SheetName.substring(0, [System.Math]::Min(30, $SheetName.Length))
            $EachWorkSheet.Name = $SheetName
        }


        $file = Get-Content $CSV
        Write-Host "    $($file.count) entries " -NoNewline
        $increment = 0

        foreach ($Line in $file) {
            
            $LineContents = $Line -split ',(?!\s*\w+")'
            foreach ($Cell in $LineContents) {
                if ($Cell[0] -eq '"' -and $Cell[-1] -eq '"') {
                    $Cell = ($Cell).trim('"')
                }
                elseif ($Cell[0] -eq "'" -and $Cell[-1] -eq "'") {
                    $Cell = ($Cell).trim("'")
                }
                $CompiledWorkSheet.Cells.Item($CompiledRow,$Column) = $Cell

                if ($IndividualTabs) {
                    $EachWorkSheet.Cells.Item($Row,$Column) = $Cell
                }

                $Column += 1

                <#
                    0  = no color
                    1  = black
                    2  = white
                    3  = red
                    4  = green
                    5  = blue
                    6  = yellow
                    7  = pink
                    8  = cyan
                    9  = bRown / dark red
                    10 = dark green
                    11 = dark blue
                    12 = olive drab
                    13 = dark purple
                    14 = turquoise
                    15 = light gray
                    16 = dark gray
                    17 = lavender
                    18 = purple / fuschia
                    19 = light yellow
                    20 = light cyan
                    21 = darker purple
                    22 = salmon
                    23 = light blue
                    24 = lilac purple
                #>
                
                if     ($LineContents[13] -match 'Critical') {
                    $excelapp.Cells.Item($CompiledRow,14).Interior.ColorIndex=9
                }
                elseif ($LineContents[13] -match 'High') {
                    $excelapp.Cells.Item($CompiledRow,14).Interior.ColorIndex=3
                }
                elseif ($LineContents[13] -match 'Medium') {
                    $excelapp.Cells.Item($CompiledRow,14).Interior.ColorIndex=22
                }
                elseif ($LineContents[13] -match 'Low') {
                    $excelapp.Cells.Item($CompiledRow,14).Interior.ColorIndex=6
                }
                elseif ($LineContents[13] -match 'Negligible') {
                    $excelapp.Cells.Item($CompiledRow,14).Interior.ColorIndex=19
                }
                elseif ($LineContents[13] -match 'Unknown') {
                    $excelapp.Cells.Item($CompiledRow,14).Interior.ColorIndex=15
                }
                elseif ($LineContents[13] -match 'Severity') {
                    $excelapp.Cells.Item($CompiledRow,14).Interior.ColorIndex=0
                }
                else {
                    $excelapp.Cells.Item($CompiledRow,14).Interior.ColorIndex=8
                }

                if ($LineContents[14] -eq '"fixed"') {
                    $excelapp.Cells.Item($CompiledRow,15).Interior.ColorIndex=4
                }
                elseif ($LineContents[14] -match 'not-fixed') {
                    $excelapp.Cells.Item($CompiledRow,15).Interior.ColorIndex=3
                }
                elseif ($LineContents[14] -match 'wont-fix') {
                    $excelapp.Cells.Item($CompiledRow,15).Interior.ColorIndex=9
                }
                elseif ($LineContents[14] -match 'unknown') {
                    $excelapp.Cells.Item($CompiledRow,15).Interior.ColorIndex=15
                }
                elseif ($LineContents[14] -match 'FixState') {
                    $excelapp.Cells.Item($CompiledRow,15).Interior.ColorIndex=0
                }
                else {
                    $excelapp.Cells.Item($CompiledRow,15).Interior.ColorIndex=8
                }
                #>
            }
            $Column = 1
            $Row += 1
            $CompiledRow += 1
            $increment += 1

            try {
                $Percent = [Math]::round(($increment / $($file.count)) * 100,0)
            }
            catch {}
            Write-Progress -Activity "Combining And Processing CSV File: $CSV" -Status "$($Percent)% Completed" -PercentComplete $Percent
        }
        Write-Host ''
    }
    Write-Host '[!] Saving to file... ' -NoNewline
    $output = $PathToCSVFiles + "\" + $OutputXlsxFileName
    $xlsx.SaveAs($output)
    Write-Host 'done'
    Write-Host "    $OutputXlsxFileName"
    $excelapp.quit()

    Invoke-Item $OutputXlsxFileName
    #$AllCVEs = Import-Csv $OutputXlsxFileName
}
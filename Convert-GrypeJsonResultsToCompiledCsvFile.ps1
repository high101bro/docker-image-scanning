Param(
    [switch]$IndividualTabs,
    [switch]$OutputExcelDoc
)
$ReportPath = Get-Location | Select-Object -ExpandProperty Path
$Reports = Get-ChildItem $ReportPath | Where-Object {$_.Extension -eq '.json'}
$script:CompiledEntryItem = 0

Foreach ($Report in $Reports) {
    Write-Host "[!] Processing JSON Report:`n    $($Report.BaseName)"
    $Contents = Get-Content $Report.FullName | ConvertFrom-Json | Select-Object -ExpandProperty matches |Select-Object -ExpandProperty vulnerability

    Write-Host "[!] Converting JSON Report To CSV Format: $($Report.BaseName)`n    $($Contents.count) entries "

    $Data = @()
    $script:EntryItem = 0
    foreach ($script:Entry in $Contents) {
        $Data += [PSCustomObject]@{
            CompiledEntryId     = "$($script:CompiledEntryItem + 1)"
            EntryId             = "$($script:EntryItem + 1)"
            Image               = "$($Report.BaseName)"
            Id                  = "$($script:Entry.Id)"
            DataSource          = "$($script:Entry.dataSource)"
            Namespace           = "$($script:Entry.Namespace)"
            URLs                = "$($script:Entry.Urls)"
            CvssVendor          = "$($script:Entry.cvss.vendorMetadata)"
            CvssVersion         = "$($script:Entry.cvss.version)"
            Vector              = "$($script:Entry.cvss.vector)"
            BaseScore           = "$($script:Entry.cvss.metrics.baseScore)"
            ExploitabilityScore = "$($script:Entry.cvss.metrics.exploitabilityScore)"
            ImpactScore         = "$($script:Entry.cvss.metrics.impactScore)"
            Severity            = "$($script:Entry.Severity)"
            FixState            = "$($script:Entry.Fix.state)"
            FixVersions         = "$($script:Entry.Fix.Versions)"
            AdvisoriesId        = "$($script:Entry.Advisories.Id)"
            AdvisoriesLink      = "$($script:Entry.Advisories.Link)"
        }
        $script:EntryItem += 1
        $script:CompiledEntryItem += 1
        try {
            $ConvertingPercent = [Math]::round(($script:EntryItem / $($Contents.count)) * 100,0)
        }
        catch {}
        Write-Progress -Activity "Converting to CSV format: $($Report.BaseName)" -Status "$($ConvertingPercent)% Completed" -PercentComplete $ConvertingPercent

    }
    $Data | Export-CSV "$ReportPath/$($Report.BaseName).csv" -NoTypeInformation -Force
}







if ($OutputExcelDoc) {
    function Combine-CSVFilesIntoSheets {
        param(
            [string]$PathToCSVFiles = '.'
        )
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
                    <#
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
        $AllCVEs = Import-Csv $OutputXlsxFileName
    }

    Combine-CSVFilesIntoSheets -PathToCSVFiles $ReportPath
}



param (
    $AllCveDetailsPath = 'C:\Users\danie\Downloads\CVE_Container_Scan_Results_2022-01-28-slim - ResidualCVEs-All.csv'
)

$AllCveDetails = Import-Csv $AllCveDetailsPath

$CveDictionary = @{}
foreach ($Cve in $AllCveDetails) {
    #if ($Cve.Id -notin $CveDictionary) {
    $CveDictionary[$Cve.Id] += "$($Cve.Image)`r`n"
    #}        
}


$GroupedCVEs = @()
foreach ($Id in ($AllCveDetails | Group-Object Id)) {
    $Data = $Id | Select-Object -ExpandProperty Group
    

    # Combines the Justification Fields
    $Justification = ''
    $Justifications = $Data | Select-Object -Property Image, Justification
    foreach ($j in $Justifications){
        if ($j.Justification) {
            $Justification += "$($j | Select-Object @{n='ImageJustification';e={$_.Image + ' -- ' + $_.Justification}} | Select-Object -ExpandProperty ImageJustification -Unique)`r`n"
        }
    }

    # Combines the Question Fields
    $Question = ''
    $Questions = $Data | Select-Object -Property Image, Question
    foreach ($q in $Questions){
        if ($q.Question) {
            $Question += "$($q | Select-Object @{n='ImageQuestion';e={$_.Image + ' -- ' + $_.Question}} | Select-Object -ExpandProperty ImageQuestion -Unique)`r`n"
        }
    }

    # Combines the Needed Fields
    $EscalationNeeded = ''
    $EscalationNeededQuestion = $Data | Select-Object -Property Image, 'Escalation Needed?'
    foreach ($e in $EscalationNeededQuestion){
        if ($e.'Escalation Needed?') {
            $EscalationNeeded += "$($e | Select-Object @{n='EscalationNeeded';e={$_.Image + ' -- ' + $_.'Escalation Needed?'}} | Select-Object -ExpandProperty EscalationNeeded -Unique)`r`n"
        }
    }
    $Id = $Data | Select-Object -ExpandProperty Id -First 1
    $GroupedCVEs += [PSCustomObject]@{
        Id                  = $Id
        Image               = $CveDictionary[$Id]
        DataSource          = $Data | Select-Object -ExpandProperty DataSource -First 1
        Namespace           = $Data | Select-Object -ExpandProperty Namespace -First 1
        URLs                = $Data | Select-Object -ExpandProperty URLs -First 1
        CvssVendor          = $Data | Select-Object -ExpandProperty CvssVendor -First 1
        CvssVersion         = $Data | Select-Object -ExpandProperty CvssVersion -First 1
        Vector              = $Data | Select-Object -ExpandProperty Vector -First 1
        BaseScore           = $Data | Select-Object -ExpandProperty BaseScore -First 1
        ExploitabilityScore = $Data | Select-Object -ExpandProperty ExploitabilityScore -First 1
        ImpactScore         = $Data | Select-Object -ExpandProperty ImpactScore -First 1
        Severity            = $Data | Select-Object -ExpandProperty Severity -First 1
        FixState            = $Data | Select-Object -ExpandProperty FixState -First 1
        FixVersions         = $Data | Select-Object -ExpandProperty FixVersions -First 1
        AdvisoriesId        = $Data | Select-Object -ExpandProperty AdvisoriesId -First 1
        AdvisoriesLink      = $Data | Select-Object -ExpandProperty AdvisoriesLink -First 1
        Justification       = $Justification
        Question            = $Question
        Needed              = $EscalationNeeded
    }
}
#$GroupedCVEs | ogv
$GroupedCVEs | Export-Csv .\Response-Merged-Together.csv -NoTypeInformation -Force
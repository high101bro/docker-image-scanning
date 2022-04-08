param (
    $ReferenceObject  = (Import-Csv '.\CVE Scan Results on 28Jan22.csv'),
    $DifferenceObject = (Import-Csv '.\CVE Scan Results on 09Feb22.csv')
)

# Creates Dictionary of unique CVEs per image for the Reference Object
$ReferenceObjectCVEs = @{}
foreach ( $CVE in $ReferenceObject ) {
    $key = $CVE.Id + ' ' + $CVE.Image
    if ( $key -notin $ReferenceObjectCVEs.keys ) {
        $ReferenceObjectCVEs[$key] += $CVE
    }
}

#Creates Dictionary of unique CVEs per image for the Difference Object
$DifferenceObjectCVEs = @{}
foreach ( $CVE in $DifferenceObject ) {
    $key = $CVE.Id + ' ' + $CVE.Image
    if ( $key -notin $DifferenceObjectCVEs.keys ) {
        $DifferenceObjectCVEs[$key] += $CVE
    }
}

#Checks to see if the CVE from the Reference Object is in the Difference Object, and if so, adds it to the Residual CVEs
$ResidualCVEs = @{}
foreach ( $CVE in $ReferenceObjectCVEs.keys ) {
    if ( $CVE -in $DifferenceObjectCVEs.keys -and $CVE -notin $ResidualCVEs.keys) {
        $ResidualCVEs[$CVE] += $ReferenceObjectCVEs[$CVE]
    }
}


$ResidualCVEs.GetEnumerator() | 
    Select-Object -ExpandProperty Value | 
    Export-Csv -Path ResidualCVEs-All.csv -NoTypeInformation -Force
$ResidualCVEs.GetEnumerator() | 
    Select-Object -ExpandProperty Value | 
    Where-Object {$_.Severity -eq 'High' -or $_.Severity -eq 'Critical' } | 
    Export-Csv -Path ResidualCVEs-High-Critical.csv -NoTypeInformation -Force









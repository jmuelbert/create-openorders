<#
.Synopsis
.Description
.EXAMPLE
.INPUTS
.OUTPUTS
.NOTES
.COMPONENT
.ROLE
.FUNCTIONALITY
#>
# New-OpenOrders.ps1
#
# A = Auftrag Nr.
# B = Hauptbereich
# C = Auftragsdatum
# D = Tage offen
# E = Deb.-Nr.
# F = Deb.-Name
# G = Verkäufer Serviceberater
# H = Arbeitswert
# I = Teile
# J = Fremdleistung
# K = Andere
# L = Gesamt
# M = Auftragswert bereits geliefert

function New-OpenOrders {

    [CmdletBinding(DefaultParameterSetName='DefaultParameterSet',
                SupportsShouldProcess=$true,
                PositionalBinding=$false,
                HelpUri='https://github.com/jmuelbert/create-openorders#help',
                ConfirmImpact='Medium')]
    Param(
        # Parameter users - users in the excelfile
        [Parameter(Mandatory=$true, 
                   HelpMessage='Path and Filename for the users.csv file',
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [ValidateNotNullOrEmpty()]
        [string]$usersFile,

        # Parameter input - the excelfile with the data
        [Parameter(Mandatory=$true,
            HelpMessage='Path and Filename for the input excel.xslx data input file')]
        [ValidateNotNullOrEmpty()]
        [string]$dataFile,

        # Parameter outputPath - Path for the resultfiles.
        [Parameter(Mandatory=$true,
                HelpMessage='The Path for the Result-Files.')]
        [ValidateNotNullOrEmpty()]
        [string]$outputPath
    )

    if (!(Test-Path $usersFile)) {
        throw "Path or File to the Users File $($usersFile) is invalid. Please supply a valid File"
    }

    if (!(Test-Path $dataFile)) {
        throw "Path or File to the Data File $($dataFile) is invalid. Please supply a valid File"
    }

    if (!(Test-Path $outputPath)) {
        throw "Path to the Output Directory $($outputPath) is invalid. Please supply a valid Path"
    }

    # Get the usernames for the Report (excelfiles)
    $beraters = import-csv $usersFile
    
    # Get the Exceltable (Data)
    $c = Import-Excel $dataFile -HeaderRow 7
    <#
     $ws = $c.Workbook.WorkSheets[0]
    $ws.Cells["C1"].Value = "TageOffen"
    $ws.Cells["F1"].Value = "Berater"
    #>

    ForEach ($berater in $beraters) {
        $fileName = $berater.Name + '.xlsx'
        $pathAndFile = $outputPath + "\" + $fileName
        $c | Select-Object 'AuftragNr.', 'Auftragsdatum', 'TageOffen', 'Deb.-Nr.', 'Deb.-Name',  'Berater', 'Arbeitswert', 'Teile', 'Fremdleistung', 'Andere', 'Gesamt', 'Auftragswert bereits geliefert' | 
        Where-Object { $_.'Berater' -like $berater.Match } | Export-Excel -AutoSize -AutoFilter -Path $pathAndFile 
    }
}

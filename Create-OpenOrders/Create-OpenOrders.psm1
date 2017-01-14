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
# Create_OpenOrders.ps1
#
# Auftrag Nr.
# Hauptbereich
# Auftragsdatum
# Tage offen
# Deb.-Nr.
# Deb.-Name
# Verk√§ufer Serviceberater
# Arbeitswert
# Teile
# Fremdleistung
# Andere
# Gesamt
# Auftragswert bereits geliefert

function Create-OpenOrders {

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

    ForEach ($berater in $beraters) {
        $fileName = '.\' + $berater.Name + '.xlsx'
        $pathAndFile = $outputPath + "\" + $fileName
        $c | select 'Auftrag Nr.', 'Auftragsdatum', 'Tage offen', 'Deb.-Nr.', 'Deb.-Name', 'Berater', 'Arbeitswert', 'Teile', 'Fremdleistung', 'Andere', 'Gesamt', 'Bereits geliefert' | 
        Where-Object { $_.'Berater' -like $berater.Match } | Export-Excel -AutoFilter -AutoSize -Path $pathAndFile
    }
}


Create-OpenOrders -usersFile berater.csv -dataFile Auftr‰ge.xlsx -outputPath out
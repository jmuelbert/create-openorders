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
    $c = Import-Excel $dataFile -WorkSheetname "OFFENE AUFTR�GE"
    <#
     $ws = $c.Workbook.WorkSheets[0]
    $ws.Cells["C1"].Value = "TageOffen"
    $ws.Cells["F1"].Value = "Berater"
    #>

    ForEach ($berater in $beraters) {
        $fileName = '.\' + $berater.Name + '.xlsx'
        $pathAndFile = $outputPath + "\" + $fileName
        $c | Select-Object 'AuftragNr.', 
                @{Name="Auftragsdatum";Expression={$_.Auftragsdatum.ToString(" dd.MMM.yyyy") -f ($_.Auftragsdatum)}}, 
                'TageOffen', 'Deb.-Nr.', 'Deb.-Name',  'Berater', 
                @{Name="Arbeitswert";Expression={'{0:N2}' -f ($_.Arbeitswert)}}, 
                @{Name="Teile";Expression={'{0:N2}' -f ($_.Teile)}}, 
                @{Name="Fremdleistung";Expression={'{0:N2}' -f ($_.Fremdleistung)}}, 
                @{Name="Andere";Expression={'{0:N2}' -f ($_.Andere)}}, 
                @{Name="Gesamt";Expression={'{0:N2}' -f ($_.Gesamt)}}, 
                @{Name="Auftragswert";Expression={'{0:N2}' -f ($_.Auftragswert)}} | 
        Where-Object { $_.'Berater' -like $berater.Match } | Export-Excel -AutoSize -FreezeTopRow -AutoFilter -Path $pathAndFile 
    }
}

Create-OpenOrders -usersFile berater.csv -dataFile Auftr�ge.xlsx  -outputPath out
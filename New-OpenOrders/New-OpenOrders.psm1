
<#
.Synopsis
   Kurzbeschreibung
.DESCRIPTION
   Lange Beschreibung
.EXAMPLE
   Beispiel für die Verwendung dieses Cmdlets
.EXAMPLE
   Ein weiteres Beispiel für die Verwendung dieses Cmdlets
.INPUTS
   Eingaben in dieses Cmdlet (falls vorhanden)
.OUTPUTS
   Ausgabe dieses Cmdlets (falls vorhanden)
.NOTES
   Allgemeine Hinweise
.COMPONENT
   Die Komponente, zu der dieses Cmdlet gehört
.ROLE
   Die Rolle, zu der dieses Cmdlet gehört
.FUNCTIONALITY
   Die Funktionalität, die dieses Cmdlet am besten beschreibt
# Create_OpenOrders.ps1
#
# A = Auftrag Nr.
# B = Hauptbereich
# C = Auftragsdatum
# D = Tage offen
# E = Deb.-Nr.
# F = Deb.-Name
# G = VerkÃ¤ufer Serviceberater
# H = Arbeitswert
# I = Teile
# J = Fremdleistung
# K = Andere
# L = Gesamt
# M = Auftragswert bereits geliefert
#>

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

    Begin
    {
    
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
    
        # Get the Exceltable (Data
        $csv = Import-Excel $dataFile -HeaderRow 7

        $xlPkg = Import-Excel $dataFile -HeaderRow 7 | Export-Excel -Path temp.xlsx -PassThru

        $ws = $xlPkg.Workbook.WorkSheets[1]

        $ws.Cells["A1"].Value = "Auftragnummer"
        $ws.Cells["B1"].Value = "Hauptbereich"
        $ws.Cells["C1"].Value = "Auftragdatum"
        $ws.Cells["D1"].Value = "Tage_offen"
        $ws.Cells["E1"].Value = "Kundennummer"
        $ws.Cells["F1"].Value = "Kundenname"
        $ws.Cells["G1"].Value = "Berater"
        $ws.Cells["H1"].Value = "Arbeit"
        $ws.Cells["I1"].Value = "Teile"
        $ws.Cells["J1"].Value = "Fremdleistung"
        $ws.Cells["K1"].Value = "Andere"
        $ws.Cells["L1"].Value = "Gesamt"
        $ws.Cells["M1"].Value = "Geliefert"
 
        $ws.Cells["C2:C200"].Style.Numberformat.Format = "dd-mm-yy"
        $ws.Cells["H2:M200"].Style.Numberformat.Format = "#,##0.00"


        $ws.Cells.AutoFitColumns()

        $xlPkg.Save()
        $xlPkg.Dispose()
    }

    Process
    {
        if ($pscmdlet.ShouldProcess("Target", "Operation"))
        {
            $c = Import-Excel temp.xlsx
    
            ForEach ($berater in $beraters) {
                $fileName = '.\' + $berater.Name + '.csv'
                $pathAndFile = $outputPath + "\" + $fileName

                $c | Select-Object 'Auftragnummer', 'Auftragdatum', 'Tage_offen', 'Kundennummer', 'Kundenname',  'Berater', 'Arbeitswert', 'Teile', 'Fremdleistung', 'Andere', 'Gesamt', 'Geliefert' | 
                Where-Object { $_.'Berater' -like $berater.Match } | Export-Csv $pathAndFile      
    
            }
        }
    }

    End
    {
        Write-Host "All done..."
    }
}


New-OpenOrders -usersFile berater.csv -dataFile Aufträge-20170129.xlsx  -outputPath out
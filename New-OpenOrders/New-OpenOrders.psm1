
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
 userfile - The CSV File for Filter the Source of usernames
 datafile - The Excel Input File
 outputPath - The Path to write the results
.OUTPUTS
 CSV Files are the results of the filter.
.NOTES
   The Format of the temporary excelfile change to:
   A1 = Auftrag Nr.                     Auftragnummer
   B1 = Hauptbereich                    Hauptbereich
   C1 = Auftragsdatum                   Auftragsdatum
   D1 = Tage offen                      Tage_offen
   E1 = Deb.-Nr.                        Kundennummer
   F1 = Deb.-Name                       Kundenname
   G1 = VerkÃ¤ufer Serviceberater       Berater
   H1 = Arbeitswert                     Arbeit
   I1 = Teile                           Teile
   J1 = Fremdleistung                   Fremdleistung
   K1 = Andere                          Andere
   L1 = Gesamt                          Gesamt
   M1 = Auftragswert bereits geliefert  Geliefert
.COMPONENT
   Import-Excel
.ROLE
   Die Rolle, zu der dieses Cmdlet gehört
.FUNCTIONALITY
   Die Funktionalität, die dieses Cmdlet am besten beschreibt
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

    <#
        Check the Input Parameters if valid
    #>
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
        $ws.Cells["C1"].Value = "Auftragsdatum"
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
 
        # Set Datum-Format for this Cells
        $ws.Cells["C2:C200"].Style.Numberformat.Format = "dd-mm-yy"
        #Set Numberformat for this area 0,00
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

                $c | Select-Object 'Auftragnummer', 'Auftragsdatum', 'Tage_offen', 'Kundennummer', 'Kundenname',  'Berater', 'Arbeitswert', 'Teile', 'Fremdleistung', 'Andere', 'Gesamt', 'Geliefert' | 
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
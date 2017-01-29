# create-openorders [powershell]
    Create from a Collection of Data extractions for salespersons.
    This will be saved in a extra excelfile for each salesperson.
## Getting Started
 - Dependencies
    - powershell 
<<<<<<< HEAD
    - powershell Excel-Import
        - The can found in the [Powershell Gallery](https://www.powershellgallery.com/)
          You can here search the Module ImportExcel or use the [link](https://www.powershellgallery.com/packages/ImportExcel/2.2.10)
    - The Excel file must contain a line with the following headings: 'Auftrag Nr.', 'Auftragsdatum', 'Tage offen', 'Deb.-Nr.', 'Deb.-Name', 'Berater', 'Arbeitswert', 'Teile', 'Fremdleistung', 'Andere', 'Gesamt', 'Bereits geliefert'
    - The title must be in the seventh row
=======
    - powershell [Excel-Import](https://github.com/dfinke/ImportExcel)

### Excel-Import
    The beste way to use this is install with powershell.
    Use the Command: `Install-Module -Name ImportExcel -Scope CurrentUser`

>>>>>>> master

## Usage
 - Create-OpenOrders -usersFile berater.csv -dataFile openorders.xlsx

## Platforms

 - Microsoft
    - The Module or Script run
 - Linux or macOS
    - The Module or Script dont load the 'epp.dll'. When your the Module or
      Script start an error will occurs 
        
## License

EUPL-V1.0 © [Jürgen Mülbert](https:/github.com/jmuelbert/create-openorders)

[Return to top](#top)

<<<<<<< HEAD

=======
>>>>>>> master

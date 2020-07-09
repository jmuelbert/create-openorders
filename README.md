# New-OpenOrders [powershell]

[![Codacy Badge](https://api.codacy.com/project/badge/Grade/2f3e8a4cd15041b58bc97b87295a3a77)](https://app.codacy.com/manual/jmuelbert/create-openorders?utm_source=github.com&utm_medium=referral&utm_content=jmuelbert/create-openorders&utm_campaign=Badge_Grade_Dashboard)

Create from a Collection of Data extractions for salespersons.
This will be saved in a extra excelfile for each salesperson.

## Getting Started
 - Dependencies
    - powershell
    - powershell Excel-Import
        - The can found in the [Powershell Gallery](https://www.powershellgallery.com/)
          You can here search the Module ImportExcel or use the [link](https://www.powershellgallery.com/packages/ImportExcel/2.2.10)
    - The Excel file must contain a line with the following headings: 'Auftrag Nr.', 'Auftragsdatum', 'Tage offen', 'Deb.-Nr.', 'Deb.-Name', 'Berater', 'Arbeitswert', 'Teile', 'Fremdleistung', 'Andere', 'Gesamt', 'Bereits geliefert'
    - The title must be in the seventh row

### Excel-Import
The beste way to use this is install with powershell.
Use the Command: `Install-Module -Name ImportExcel -Scope CurrentUser`

## Usage
 - `New-OpenOrders -usersFile berater.csv -dataFile orders.xlsx`

## Platforms

 - Microsoft
    - The Module or Script run
 - Linux or macOS
    - The Module or Script dont load the 'epp.dll'. When your the Module or
      Script start an error will occurs 
        
## License

EUPL-V1.1 © [Jürgen Mülbert](https:/github.com/jmuelbert/create-openorders)

[Return to top](#top)

[Help](#help)
# Help for New-OpenOrders



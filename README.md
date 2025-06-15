# Export Data to Excel

This PowerShell script connects to a SQL Server database, executes a query, exports the result to a CSV file, and converts it into an XLSX file with two sheets: the main data and a reference sheet with static information.

## Requirements

- PowerShell 5+
- [ImportExcel PowerShell module](https://www.powershellgallery.com/packages/ImportExcel)

To install the required module:

```powershell
Install-Module -Name ImportExcel -Scope CurrentUser -Force
```

## Configuration

This script uses the following environment variables for the database connection:

- `DB_SERVER` – SQL Server address or hostname  
- `DB_NAME` – Database name  
- `DB_USER` – Username  
- `DB_PASSWORD` – Password  

Set them in your environment or within your PowerShell session before execution:

```powershell
$env:DB_SERVER = "your_server"
$env:DB_NAME = "your_database"
$env:DB_USER = "your_user"
$env:DB_PASSWORD = "your_password"
```

## Usage

Run the script using PowerShell:

```powershell
.\export_data_to_excel.ps1
```

Or run the batch file:

```bat
run_export.bat
```

## Output

- A timestamped `.xlsx` file with the results of the query.
- Sheet 1: `Data` – Results from the database.
- Sheet 2: `Reference` – Static example product info.

## Notes

- The SQL query is just a placeholder. Modify it to match your database structure.
- The script deletes the temporary CSV file after generating the XLSX.
- The script automatically installs the required PowerShell module if it is not found.

---

Feel free to customize this project for your own reporting needs.

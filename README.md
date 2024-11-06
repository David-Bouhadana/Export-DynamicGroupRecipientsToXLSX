# Export-DynamicGroupRecipientsToXLSX
This PowerShell script allows you to export the recipients of a dynamic distribution group in XLSX format.

# Export Dynamic Distribution Group Recipients to XLSX

This PowerShell script allows you to export the recipients of a dynamic distribution group in XLSX format using the `ImportExcel` module.

## Prerequisites

- PowerShell 5.1 or higher
- `ImportExcel` module

## Installing the ImportExcel Module

If the `ImportExcel` module is not already installed, you can install it by running the following command:

```powershell
Install-Module -Name ImportExcel -Scope CurrentUser
```
## Usage

Clone this repository or download the PowerShell script.

Open PowerShell and navigate to the directory containing the script.

Run the script, replacing "NameOfYourGroup" with the name of your dynamic distribution group.
# Get the dynamic distribution group
$dynamicGroup = Get-DynamicDistributionGroup -Identity "NomDeVotreGroupe"

# Get the recipients based on the recipient preview filter
$recipients = Get-Recipient -RecipientPreviewFilter $dynamicGroup.RecipientFilter -OrganizationalUnit $dynamicGroup.RecipientContainer

# Export the recipients to XLSX format
$recipients | Export-Excel -Path "$PWD\fichier.xlsx" -WorksheetName "Destinataires"

## Script Explanation

`Get-DynamicDistributionGroup -Identity "NomDeVotreGroupe"`: This command retrieves the specified dynamic distribution group.

`Get-Recipient -RecipientPreviewFilter $dynamicGroup.RecipientFilter -OrganizationalUnit $dynamicGroup.RecipientContainer`: This command retrieves the recipients based on the group's recipient preview filter.

`Export-Excel -Path "$PWD\fichier.xlsx" -WorksheetName "Destinataires"`: This command exports the recipients to an XLSX file named `fichier.xlsx` in the current working directory.

## Author

Script written by **David Bouhadana**.

- Blog: [M365 journey](https://m365journey.blog/)

## License

This project is licensed under **GNU GPL 3**. You are free to use, modify, and distribute this code as long as the modifications and derived versions are also licensed under GNU GPL 3. For more information, please refer to the full license text [GNU GPL 3](https://www.gnu.org/licenses/gpl-3.0.html).

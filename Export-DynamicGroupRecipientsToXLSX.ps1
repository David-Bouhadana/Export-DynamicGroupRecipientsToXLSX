# Obtenir le groupe de distribution dynamique
# Get dynamic distribution group
$dynamicGroup = Get-DynamicDistributionGroup -Identity "NameOfYourGroup"

# Obtenir les destinataires selon le filtre de prévisualisation des destinataires
# Get recipients according to recipient preview filter
$recipients = Get-Recipient -RecipientPreviewFilter $dynamicGroup.RecipientFilter -OrganizationalUnit $dynamicGroup.RecipientContainer

# Exporter les destinataires en format XLSX
# Export recipients in XLSX format
$recipients | Export-Excel -Path "$PWD\fichier.xlsx" -WorksheetName "Destinataires"
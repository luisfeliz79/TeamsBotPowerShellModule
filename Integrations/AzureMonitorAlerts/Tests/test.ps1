
$RelativePath='./Integrations/AzureMonitorAlerts/'

$code="function Test() {$((gc $RelativePath/code.ps1 -raw) -replace 'using namespace System.Net') }"
$Functioncode=[scriptblock]::Create($($code))
. $Functioncode

$env:TeamsBot_WebhookUrl=((az keyvault secret show --name TeamsBotWebHookUrl --vault-name felizlabs-keyvault-hub) | ConvertFrom-Json).value

Get-ChildItem -Path $RelativePath/Tests/*.json | ForEach-Object {

    Test -Request (gc $_.FullName -Raw | ConvertFrom-Json) -WhatIf


}

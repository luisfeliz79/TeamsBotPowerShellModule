using namespace System.Net


# Input bindings are passed in via param block.
param($Request, $TriggerMetadata)

# Write to the Azure Functions log stream.
Write-Host "PowerShell HTTP trigger function processed a request."

Import-Module TeamsBot -verbose

Get-module

# Grab the request sent
#$body=$Request|convertto-json -depth 99
$Body=$Request.Body

if ($Body.data.essentials.alertRule -eq $null -or $Body.data.essentials.alertRule -eq "") {

    Write-Information "This does not look like a proper request, exiting..."
    $body | convertto-json -depth 99
    $body
    exit
}


# Main code

$Card = New-TeamsBotCard

# Code for Header

switch ($Body.data.essentials.MonitorCondition) {
    'Fired' {$Style="attention"}
    'Resolved' {$Style="good"}
    Default {$Style="attention"}
}

$Columns=@()
$Columns+=@{
        type= "Column"
        width= "70"
        height="stretch"
        items=@(@{            
                type= "TextBlock"
                text= $("ALERT {0}" -f $Body.data.essentials.monitorCondition.ToUpper())
                wrap= "false"
                size= "Large"
                weight= "Bolder"
            })
        style=$Style
}
$Columns+=@{
    type= "Column"
    width= "30"
    height="stretch"
    items=@(@{            
            type= "TextBlock"
            text= $("{0}" -f $Body.data.essentials.severity.toUpper())
            wrap= "false"
            size= "Large"
            weight= "Bolder"
        })
    style="warning"       
}

$Header=@{
    type="ColumnSet"
    columns=$Columns
    #style=$Style
}


$Card.AddCustomBodyPart($Header)



# Code for Title and App Name
    $Title="{0}" -f $Body.data.essentials.alertRule
    $AppName="Azure Monitor Alerts"
    
    if ($Body.data.essentials.alertTargetIDs) {
        $ConfigItems=$Body.data.essentials.alertTargetIDs
        
        If ($ConfigItems.count -gt 1) {
            $ConfigItems[0]+=" and {0} others" -f $ConfigItems.count
        }

        $SubTitle="On {0} an alert was sent for resource`n- {1}" -f  $Body.data.essentials.firedDateTime,$ConfigItems[0]
    } else {
        $SubTitle="Alert Timestamp: {0}" -f  $Body.data.essentials.firedDateTime
    }

    $Card.AddTitle($Title)
    $Card.AddAppName($AppName)
    $Card.AddSubTitle($SubTitle)

# Code for Alert body, which depends on what details were sent

## Service Health Alert
$HTMLtoMarkDown={

    function HTMLtoMarkDown () {

    }

}

if ($Body.data.essentials.monitoringService -eq 'ServiceHealth'){

    if ($Body.data.alertContext.properties.defaultLanguageContent) {
        $Card.AddSubTitle($Body.data.alertContext.properties.title)
        $Message=$Body.data.alertContext.properties.defaultLanguageContent 
        $Card.AddTextBlock($Message)
    }

    if ($Body.data.alertContext.properties.trackingId) {
       $ServiceHealthLink="https://portal.azure.com/#blade/Microsoft_Azure_Health/AzureHealthBrowseBlade/trackingId/{0}" -f $Body.data.alertContext.properties.trackingId
       $Card.AddAction("View Service Health",$ServiceHealthLink)
    }
 

} # End Service Health Alert


## Others
if ($Body.data.alertContext.properties.statusMessage) {

    $Card.AddSubTitle("Additional Details`n")
    $Message="*StatusMessage*: {0}" -f $Body.data.alertContext.properties.statusMessage
    $Card.AddTextBlock($Message)
}

if ($Body.data.alertContext.condition.allOf.dimensions) {

    $Card.AddSubTitle("Additional Details`n")
  
    # If Dimensions were included, show them in a table
    $MyTable=@()
    $MyObject=New-Object PSCustomObject 
    $Body.data.alertContext.condition.allOf.dimensions | ForEach-Object {
        $name=$_.Name
        $value=$_.value
        #Write-Warning "$name => $value"
              
        $MyObject | Add-Member -MemberType NoteProperty -Name $name -Value $value 
        
    }
    $MyTable+=@($MyObject)

    $Card.AddTable($MyTable)
} # end others

# Code for Action buttons
if ($Body.data.alertContext.condition.allOf.linkToFilteredSearchResultsUI) {
    $Card.AddAction("View Search Results",$Body.data.alertContext.condition.allOf.linkToFilteredSearchResultsUI)
    $CalcPortalLink="https://{0}/resource{1}" -f ($Body.data.alertContext.condition.allOf.linkToFilteredSearchResultsUI -split '/')[2],$Body.data.essentials.alertId
    $Card.AddAction("View Resource",$CalcPortalLink)
}
#$Card.AddAction("View Alerts","https://portal.azure.com/#view/Microsoft_Azure_Monitoring/AzureMonitoringBrowseBlade/~/alertsV2")
$AlertUrl="https://portal.azure.com/#blade/Microsoft_Azure_Monitoring/AlertDetailsTemplateBlade/alertId/{0}/invokedFrom/emailcommonschema" -f [System.Web.HttpUtility]::UrlEncode($Body.data.essentials.alertId)
$Card.AddAction("View Alert",$AlertUrl)


Send-TeamsBotMessage -verbose -Card $Card.GetCard() 
Write-Information "Sent Message"

# Associate values to output bindings by calling 'Push-OutputBinding'.
Push-OutputBinding -Name Response -Value ([HttpResponseContext]@{
    StatusCode = [HttpStatusCode]::OK
    #Body = $body
})



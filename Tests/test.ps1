# Function Test() {,">>> Failed <<<"
#     try {
#          TestResult -Name "Test"-Status "Passed"
#      } catch {
#          TestResult -Name "Test"-Status ">>> Failed <<<"
#      }
#  }  


function TestResult($name,$status) {

[PSCustomObject]@{
    TestName = $Name
    Result   = $Status
}

}
  
Function TestCustomCard () {

    try {
    # Use a customized card and send the message

        # Create a Card and add title, and optionally subtitle, and AppName
        $SampleCard=New-TeamsBotCard
        $SampleCard.AddTitle("Test Title")  
        $SampleCard.AddAppName("[Created with TeamBot Powershell module]($GitHubRepo)")                                                                                                                                                                     
        $SampleCard.AddSubTitle("Test SubTitle")                                                                                                                                                                                                                                                                                                                                                   

        # Add a Textblock
        $SampleCard.AddTextBlock(@"

        *This is the main textblock which supports Markdown*

        - section 1
        - section 2
        - section 3

"@)   # this line needs to be flushed on the left

        
        $myTable=@()
        $myTable+=[PsCustomObject]@{Resource="Graph";App="Windows";User="Luis"}
        $myTable+=[PsCustomObject]@{Resource="Graph";App="Office";User="Mike"}
        $SampleCard.AddTable($myTable)



        # Add action buttons, multiple is Ok                                                                                                                                                
        $SampleCard.AddAction("Action 1","https://bing.com")
        $SampleCard.AddAction("Action 2","https://bing.com")

        # Send the card
        Send-TeamsBotMessage -Card $SampleCard -WhatIf -WebhookURL "https://bing.com"
        TestResult -Name "TestCustomCard"-Status "Passed"
    } catch {
        TestResult -Name "TestCustomCard"-Status ">>> Failed <<<"
    }

}

Function TestSimpleCards() {
   try {
        Send-TeamsBotMessage -Message "Message" -Title "Title" -WhatIf -WebhookURL "https://bing.com"
        Send-TeamsBotMessage -Message "Message" -Title "Title" -SubTitle "SubTitle" -WhatIf -WebhookURL "https://bing.com"
        Send-TeamsBotMessage -Message "Message" -Title "Title" -SubTitle "SubTitle" -AppName "AppName" -WhatIf -WebhookURL "https://bing.com"

        $myTable=@()
        $myTable+=[PsCustomObject]@{Resource="Graph";App="Windows";User="Luis"}
        $myTable+=[PsCustomObject]@{Resource="Graph";App="Office";User="Mike"}
        Send-TeamsBotMessage -Message "Message" -Title "Title" -SubTitle "SubTitle" -AppName "AppName" -Table $myTable  -WebhookURL "https://bing.com"
        
        TestResult -Name "TestSimpleCards"-Status "Passed"

    } catch {
        TestResult -Name "TestSimpleCards"-Status ">>> Failed <<<"
    }
}



Function TestSetWebhookUrl() {
    try {
         Set-TeamsBotWebhookUrl -WebhookURL "https://bing.com" -Verbose
         TestResult -Name "TestSetWebhookUrl"-Status "Passed"
     } catch {
         TestResult -Name "TestSetWebhookUrl"-Status ">>> Failed <<<"
     }
 }

Function TestClearWebhookUrl() {
    try {
         Clear-TeamsBotWebhookUrl -Verbose
         TestResult -Name "TestClearWebhookUrl"-Status "Passed"
     } catch {
         TestResult -Name "TestClearWebhookUrl"-Status ">>> Failed <<<"
     }
 }

Function TestSendWithStoredWebhookUrl() {
    try {
        Set-TeamsBotWebhookUrl -WebhookURL "https://bing.com"
        Send-TeamsBotMessage -Message "Message" -Title "Title" -WhatIf -Verbose
        Clear-TeamsBotWebhookUrl
        TestResult -Name "TestSendWithStoredWebhookUrl"-Status "Passed"
     } catch {
         TestResult -Name "TestSendWithStoredWebhookUrl"-Status ">>> Failed <<<"
     }
 }

 Function TestSendWithEnvWebHookUrl() {
    try {
        Clear-TeamsBotWebhookUrl
        $env:TeamsBot_WebhookUrl="https://bing.com"
        Send-TeamsBotMessage -Message "Message" -Title "Title" -WhatIf -Verbose
        TestResult -Name "TestSendWithEnvWebHookUrl"-Status "Passed"
     } catch {
         TestResult -Name "TestSendWithEnvWebHookUrl"-Status ">>> Failed <<<"
     }
 }
 
 


$ErrorActionPreference="Stop"

$Version="1.0.3"

$Testing=@()
$Error.clear()
Test-ModuleManifest -path ".\TeamsBot\$version\TeamsBot.psd1"
Import-Module .\TeamsBot\$version\teamsbot.psd1 -Force

$Testing+=TestSimpleCards
$Testing+=TestCustomCard
$Testing+=TestSetWebhookUrl
$Testing+=TestClearWebhookUrl
$Testing+=TestSendWithStoredWebhookUrl
$Testing+=TestSendWithEnvWebHookUrl

Remove-Module TeamsBot


$Testing | ft
$error


Import-module PSScriptAnalyzer
Invoke-ScriptAnalyzer .\TeamsBot\$version -ExcludeRule PSAvoidTrailingWhitespace,PSAvoidUsingConvertToSecureStringWithPlainText 

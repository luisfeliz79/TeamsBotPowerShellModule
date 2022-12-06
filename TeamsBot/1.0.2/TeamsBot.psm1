

function Send-TeamsBotMessage () {
  <#
   .Synopsis
    Send messages to a Teams Channel using Powershell!
  
   .Description
    This module allows you to send messages to Teams channels using the Webhooks Feature of Microsoft teams.
    It supports sending customized cards for more advanced scenarios.
  
  
             
    .Parameter Message
    The message to be sent to the Teams channel
    Some markdown features are possible, see here:
    https://support.microsoft.com/en-us/office/use-markdown-formatting-in-teams-4d10bd65-55e2-4b2d-a1f3-2bebdcd2c772
  
    .Parameter Title
    A title to be included with the message
  
    .Parameter AppName
    Optionally include the app name and make it clickiable using markdown
  
    .Parameter SubTitle
    Optionally include a Subtitle which will appear bolded
  
    .Parameter Table
    Optionally include a table of Key/value pairs (Array of PsCustomObject)
    Ex. an Object formatted like this :
    $myTable=@()
    $myTable+=[PsCustomObject]@{Resource="Graph";App="Windows";User="Luis"}
    $myTable+=[PsCustomObject]@{Resource="Graph";App="Office";User="Mike"}
  
    .Parameter ActionButtonTitle
    Optionally include a button with this title
  
    .Parameter ActionButtonLink
    Link for the button
  
  
    .Parameter Card
    For more advanced scenarios, a custom card object can be created and then passed with this parameters
    See $GitHubLink for more information
  
    .Parameter WebhookURL
    The WebhookURL to send the message to.
    The Url can also be persisted using Set-TeamsBotWebhookUrl or specified using Environment variable TeamsBot_WebHookUrl
  
   .Example
     # Preconfigure the WebhookUrl
     Set-TeamsBotWebhookUrl -WebhookURL 'https://url'
  
   .Example
     # Send a basic message
     Send-TeamsBotMessage -Message "Hello" -Title "My Title" -WebhookUrl
  
   .Example
     # Send a basic message including the WebhookUrl inline
     Send-TeamsBotMessage -Message "Hello" -Title "My Title" -WebhookUrl 'https://url'
  
   .Example
     # Send a basic message and include an action button
     Send-TeamsBotMessage -Message "Hello" -Title "My Title" -ActionButtonTitle "View site" -ActionButtonLink "https://bing.com"
  
   .Example
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
  
  "@)
  
      # Optionally add a table to neatly display Key/Value pairs
      $myTable=@()
      $myTable+=[PsCustomObject]@{Resource="Graph";App="Windows";User="Luis"}
      $myTable+=[PsCustomObject]@{Resource="Graph";App="Office";User="Mike"}
  
  
  
      # Add action buttons, multiple is Ok                                                                                                                                                
      $SampleCard.AddAction("Action 1","https://bing.com")
      $SampleCard.AddAction("Action 2","https://bing.com")
  
      # Send the card
      Send-TeamsBotMessage -Card $SampleCard
  
  #>

  [CmdletBinding(SupportsShouldProcess)]
   param(

      [Parameter(Mandatory = $true, ParameterSetName = 'Directmessage')]
                  
      [String]$Message,
  
      [Parameter(Mandatory = $true, ParameterSetName = 'Directmessage')] 
      [String]$Title,
  
      [Parameter(Mandatory = $false, ParameterSetName = 'Directmessage')] 
      [String]$AppName,
  
      [Parameter(Mandatory = $false, ParameterSetName = 'Directmessage')] 
      [String]$SubTitle,
  
      [Parameter(Mandatory = $false, ParameterSetName = 'Directmessage')] 
      $Table,
  
      [Parameter(Mandatory = $false, ParameterSetName = 'Directmessage')]
      [Parameter(Mandatory = $false, ParameterSetName = 'IncludeAction')] 
      [String]$ActionButtonTitle="View Details",
  
      [Parameter(Mandatory = $true, ParameterSetName = 'IncludeAction')] 
      [Parameter(Mandatory = $false, ParameterSetName = 'Directmessage')] 
      [String]$ActionButtonLink,
  
      [Parameter(Mandatory = $true, ParameterSetName = 'UsingCard')] 
      [PsCustomObject]$Card,

      [Parameter(Mandatory = $true, ParameterSetName = 'UsingCardJson')] 
      [PsCustomObject]$CardJson,
  
      [Parameter(Mandatory = $false)] 
      [String]$WebhookURL

      #[Parameter(Mandatory = $false)] 
      #[Switch]$WhatIf
 
  
    )
  
  
  
  
    if ($Message) {
  
          # Lets use the default template and card setup
  
          Write-Verbose -Message "Using Default template"
  
          $DefaultCard=New-TeamsBotCard
          $DefaultCard.AddTitle($Title)
          
          if ($null -ne $AppName) { $DefaultCard.AddAppName($AppName)}
          if ($null -ne $SubTitle) { $DefaultCard.AddSubTitle($SubTitle)}                                                                                                                                                                   
                                                                                                                              
          $DefaultCard.AddTextBlock($Message)
          
          if ($null -ne $ActionButtonTitle -and $ActionButtonLink ) { 
                  $DefaultCard.AddAction($ActionButtonTitle,$ActionButtonLink) 
          }
          
          if ($null -ne $Table) {$DefaultCard.AddTable($Table)}
              
          try {
              $UseThisCard = $DefaultCard.GetCard()
          } catch {
              Write-Error "The card could not be rendered, check the template file $CustomTemplatePath -or if you suspect a bug, please file an issue at $GitHubRepo"
          }
    }
  
    if ($Card) {
  
      if ($Card.type -eq 'AdaptiveCard') {
  
        Write-Verbose -Message "Using a pre created Card"
  
        # This will take in a Card object directly and use it.
  
        try {
            $UseThisCard = $Card.GetCard()
        } catch {
            Write-Error "The card could not be rendered, for assistance or to file an issue, please visit $GitHubRepo"
        }
      } else {
        Write-Error "The supplied Card does not appear to be valid"
        break
      }
  
    }

    If ($CardJson) {
      # A Raw card in JSON Format
      $CardObject=$CardJson | ConvertFrom-Json

      if ($CardObject.type -eq 'AdaptiveCard') {
  
        Write-Verbose -Message "Using a pre created Card in JSON format"
  
        # This will take in a Card object directly and use it.
  
        try {
            $UseThisCard = $CardObject
        } catch {
            Write-Error "The card could not be rendered, for assistance or to file an issue, please visit $GitHubRepo"
        }
      } else {
        Write-Error "The supplied Card does not appear to be valid"
        break
      }

    }
  
    if ($WebhookURL) {
      
      Write-Verbose -Message "Using WebhooUrl from command line parameter"
  
  
    } else {
      # If webhook was not specified then
      # check for other sources
      Write-Verbose -Message "Obtaining configured WebHookUrl"
      $WebHookURL = Get-TeamsBotWebhookUrl
    }
  
  
    # If we got this far, lets send the message
  
    # Prepare the complete payload, and convert to JSON format
    $jsonPayload=@{
          type="message"
          attachments=@(@{
                contentType="application/vnd.microsoft.card.adaptive"
                contentUrl="https://github.com/luisfeliz79/TeamBotPowerShellModule"
                content=$UseThisCard
              })
    } | ConvertTo-Json -Depth 99
  
  
    if ($null -ne $jsonPayload) {
      Write-Verbose "Payload:"
      Write-Verbose $jsonPayload
      
      If ($WhatIfPreference -eq $true) {
        Write-Warning "WhatIf: Message would had been sent."  
        
      } else {
        
          Write-Verbose "Sending message...."
          $UserAgent="TeamBotPowerShellModule-$(($MyInvocation.MyCommand.Version).tostring())"        
          $response=Invoke-RestMethod -Method Post -Uri $WebHookURL -ContentType 'application/json' -Body $jsonPayload -UserAgent $UserAgent
          Write-Verbose ($response | ConvertTo-Json -Depth 99)
      }
    }
  
  
  
  
  }
  
  function New-TeamsBotCard() {
  
  <#
   .Synopsis
    Functions to create an Adaptive card. for Help see "https://github.com/luisfeliz79/TeamsBotPowerShellModule"
    Some markdown features are possible, see here:
    https://support.microsoft.com/en-us/office/use-markdown-formatting-in-teams-4d10bd65-55e2-4b2d-a1f3-2bebdcd2c772
  
   .Example
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
  
  "@)
  
      # Optionally add a table
      $myTable=@()
      $myTable+=[PsCustomObject]@{Resource="Graph";App="Windows";User="Luis"}
      $myTable+=[PsCustomObject]@{Resource="Graph";App="Office";User="Mike"}
      $SampleCard.AddTable($myTable)
  
  
  
      # Add action buttons, multiple is Ok                                                                                                                                                
      $SampleCard.AddAction("Action 1","https://bing.com")
      $SampleCard.AddAction("Action 2","https://bing.com")
  
      # Send the card
      Send-TeamsBotMessage -Card $SampleCard
  #>
  
  [CmdletBinding(SupportsShouldProcess)]
  param()
  
  # Define some functions for this object type
  $AddTextBlock={
  param ([string]$value)
      $This.body+=( @{
          type="TextBlock"
          text=HtmlToMarkDown -html $value
          isSubtle="true"
          wrap="true"
      } )
  }
  
  $AddSubTitle={
  param ([string]$value)
      $This.body+=( @{
          type="TextBlock"
          text=$value
          isSubtle="true"
          wrap="true"
          weight='Bolder'
      } )
  }
  
  $AddTitle={
  param ([string]$value)
      $This.body+=( @{
          type="TextBlock"
          text=$value
          isSubtle="true"
          wrap="true"
          size='Large'
          weight='Bolder'
      } )
  }
  
  $AddAppName={
  param ([string]$value)
      $This.body+=( @{
          type="TextBlock"
          text=$value
          isSubtle="true"
          color='Accent'
          weight='Bolder'
          size='Small'
          spacing='None'
      } )
  }
  
  $AddCustomBodyPart={
  param ($value)
      $This.body+=( $value )
  }
  
  
  
  $AddAction={
  param ([string]$value1,[string]$value2)
      $This.actions+=( 
      @{
        type='Action.OpenUrl'
        title=$value1
        url=$value2
      }
      )
  }
  
  $AddTable={
    param ($items)
  
      $RowsObject=@()
      $ColumnObjects=@()
      

      if ($items.gettype().Name -eq 'Object[]' -and $items[0].gettype().Name -eq 'PSCustomObject'  ) {

        # Gets the list of properties
        $Props=($items | Get-Member | Where-Object membertype -eq "NoteProperty").name 
        Write-Verbose "Found Properties $($Props -join ', ')"
        # Title Row
        $FirstRow=$Props | ForEach-Object {
          @{
                type="TableCell"
                items=@(
                    @{
                        type="TextBlock"
                        text=$_
                        isSubtle="true"
                        wrap="true"
                        weight="Bolder"
                    })
          }
        }

        # Creates the first row with titles based on PSCustomObject NoteProperties
        $RowsObject+=(
          @{
              type="TableRow"
              cells=$FirstRow
            }
        )

        # Creates the columns definitions
        $ColumnObjects+=$Props | ForEach-Object {
          @{ width=1 }
        }
      } else {
        Write-error "Object not supported, use an array of PSCustomObject [PSCustomObject]@{Prop1='Value1';Prop2='Value2'}"
        break
      }

      # Creates the data rows
      $items | Foreach-Object {
        $Cells=@()
        $curItem=$_
        $Cells+=$Props | ForEach-Object {
          @{
            type="TableCell"
                items=@(
                    @{
                        type="TextBlock"
                        text=$curItem.$_
                        isSubtle="true"
                        wrap="true"
                        
                    })
          }
        }
       
          $RowsObject+=(
            @{
                type="TableRow"
                cells=$Cells
              }
            )
      }
  
      # Create the overarching Table object
      $TableObject=[PsCustomObject]@{
      
              type="Table"
              columns=@($ColumnObjects)  
              rows=@($RowsObject)
              showGridLines="true"
      
      }
  
      # Add it to the existing body
      $This.body+=($TableObject)
      
  }
      
  
  
  
  
  $GetCard={
      return $This
  }
  
  # Create the high level object
  $NewTeamsBotCard=[PsCustomObject]@{
          '$schema'="http://adaptivecards.io/schemas/adaptive-card.json"
          type="AdaptiveCard"
          version='1.4'
          body=@()
          actions= @()
      }
  # Add functions
  $NewTeamsBotCard | Add-member -MemberType ScriptMethod -Name AddTextBlock -Value $AddTextBlock -Force
  $NewTeamsBotCard | Add-member -MemberType ScriptMethod -Name AddSubTitle -Value $AddSubTitle -Force
  $NewTeamsBotCard | Add-member -MemberType ScriptMethod -Name AddTitle -Value $AddTitle -Force
  $NewTeamsBotCard | Add-member -MemberType ScriptMethod -Name AddAppName -Value $AddAppName -Force
  $NewTeamsBotCard | Add-member -MemberType ScriptMethod -Name AddAction -Value $AddAction -Force
  $NewTeamsBotCard | Add-member -MemberType ScriptMethod -Name AddTable -Value $AddTable -Force
  $NewTeamsBotCard | Add-member -MemberType ScriptMethod -Name AddCustomBodyPart -Value $AddCustomBodyPart -Force
  $NewTeamsBotCard | Add-member -MemberType ScriptMethod -Name GetCard -Value $GetCard -Force
  
  return $NewTeamsBotCard 
  
   
  }
  function Set-TeamsBotWebhookUrl () {
  <#
   .Synopsis
    Persists the WebHookUrl securely to an encrypted file - $HOME/TeamsBotEncryptedStorage.enc
    After this step has been completed, the WebHookUrl does not have to specified on Send-TeamsBotMessage commands
  #>
  [CmdletBinding(SupportsShouldProcess)]
   param(
      [Parameter(Mandatory = $true)]
      [String]$WebhookURL
    )
  
    try {
      Write-Verbose  "Saving WebhookURL as a SecureString to $HOME/TeamsBotEncryptedStorage.enc ..."
      $WebhookURL| ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString | Out-File "$HOME/TeamsBotEncryptedStorage.enc" -ErrorAction Stop -WhatIf:$WhatIfPreference
    } catch {
      Write-Error "Unable to save WebhookURL To file $HOME/TeamsBotEncryptedStorage.enc"
    }
  
  }
  function Clear-TeamsBotWebhookUrl () {
  <#
   .Synopsis
    Clears the WebHookUrl by deleting file $HOME/TeamsBotEncryptedStorage.enc
   
  #>
  [CmdletBinding(SupportsShouldProcess)]
  param()

     try {
      Write-Verbose  "Removing file $HOME/TeamsBotEncryptedStorage.enc ..."
      if ((Test-Path -Path "$HOME/TeamsBotEncryptedStorage.enc") -eq $true){
          Remove-item -Path "$HOME/TeamsBotEncryptedStorage.enc" -WhatIf:$WhatIfPreference
        } else {
          Write-Verbose  "File not found, nothing to do."
        }
    } catch {
      Write-Error "Unable to remove file $HOME/TeamsBotEncryptedStorage.enc"
    }
  }
  function Get-TeamsBotWebhookUrl () {
  
  <#
   .Synopsis
    Gets the WebHookUrl either from Environment variable TeamsBot_WebHookUrl (first) or if it exists, from $HOME/TeamsBotEncryptedStorage.enc (second)
  #>
  
  # This function will try several methods to obtain the WebhookURL
  
  # First try checking for an Environment variable
  
  
      if ($Env:TeamsBot_WebHookUrl) {
  
          Write-Verbose "Using WebhookUrl found in TeamsBot_WebHookUrl environment variable"
          return $Env:TeamsBot_WebHookUrl
  
      }
  
      if ((Test-Path -Path $HOME/TeamsBotEncryptedStorage.enc) -eq $true) {
  
          Write-Verbose "Using WebHookURL found in $HOME/TeamsBotEncryptedStorage.enc"
  
          try {
  
              $BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR((Get-Content -Path "$HOME/TeamsBotEncryptedStorage.enc" | ConvertTo-SecureString))
              $URL=[System.Runtime.InteropServices.Marshal]::PtrToStringUni($BSTR)
              return $URL
          } catch {
              Write-Error "Could not read from $HOME/TeamsBotEncryptedStorage.enc"
              Exit
          }
  
      }
  
      Write-error "WebhookURL has not been configured, use Set-TeamsBotWebhookUrl, or environment variable TeamsBot_WebHookUrl"
      break
  }

  function HTMLtoMarkDown ($html) {

    $LinkTranslation=@()
    
    $anchorArr=$html -split '<a'

    $anchorArr| ForEach-Object {

        if ($_ -like '*</a>*') {
            $HtmlLinkDef='<a'+($_ -split '</a>')[0]+'</a>'

            $hrefArr=$HtmlLinkDef -split 'href='
            $Link=($hrefArr[1] -split ' |>')[0] -replace '"'
            $Caption=(($hrefArr[1] -split '>')[1] -split '</a')[0]

            $LinkTranslation+=@{
                ReplaceThis=$HtmlLinkDef
                WithThis="[$Caption]($Link)"
            }

        }


    }
    #return $LinkTranslation
    $LinkTranslation | ForEach-Object {
        $html=$html -ireplace [regex]::Escape($_.ReplaceThis), $_.WithThis
    }

    $html = $html -replace '<p>|</p>' `
                  -replace '<strong>|</strong>','**' `
                  -replace '<br>',"`r`n" `
                  -replace '&nbsp;' `
                  -replace ' rel="noopener noreferrer" target="_blank"',''

    return $html
}

  
  $GitHubRepo="https://github.com/luisfeliz79/TeamsBotPowerShellModule"
  
  
  $ErrorActionPreference="Stop"
  # Export-ModuleMember -Cmdlet * -Function *
  
  Export-ModuleMember -Function Send-TeamsBotMessage
  Export-ModuleMember -Function Set-TeamsBotWebhookUrl
  Export-ModuleMember -Function Clear-TeamsBotWebhookUrl
  Export-ModuleMember -Function New-TeamsBotCard
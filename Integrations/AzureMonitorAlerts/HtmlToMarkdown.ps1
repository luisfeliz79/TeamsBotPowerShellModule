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
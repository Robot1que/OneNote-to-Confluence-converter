function ConvertTo-Page
{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [PSObject]$InputObject,
        [Parameter(Mandatory = $true)]
        [PSObject]$SpaceKey,
        [Parameter(Mandatory = $false)]
        [PSObject]$ParentPageId
    )
    Begin
    {
        . (Join-Path $PSScriptRoot .\New-Page.ps1)
    }
    Process
    {
        $title = $InputObject.title.Trim()
        
        if ($Global:pageNameTitle.ContainsKey($title))
        {
            $Global:pageNameTitle[$title] = $Global:pageNameTitle[$title] + 1
        }
        else
        {
            $Global:pageNameTitle.Add($title, 1)
        }

        $duplicatePageNum = $Global:pageNameTitle[$title]
        if ($duplicatePageNum -gt 1)
        {
            $title += " ($duplicatePageNum)"
        }

        Write-Verbose "Creating '$title' page for OneNote $($InputObject.type) '$($InputObject.title)'..."
        Write-Verbose "pageNameTitle count: $($Global:pageNameTitle.Count)"

        [string]$url = [Security.SecurityElement]::Escape($InputObject.url)
        $url = $url.Insert(8, "///")

        [string]$escapedTitle = [Security.SecurityElement]::Escape($InputObject.title)

        $newPage = `
            New-Page `
                -Title $title `
                -SpaceKey $SpaceKey `
                -ParentPageId $ParentPageId `
                -Content "<a href=`"$url`">$escapedTitle</a>" `
                -Credential $credential

        if ($InputObject.type -eq "SectionGroup")
        {
            foreach($section in $InputObject.sections)
            {
                ConvertTo-Page `
                    -InputObject $section `
                    -SpaceKey $SpaceKey `
                    -ParentPageId $newPage.id
            }
        }

        if ($InputObject.type -eq "Section")
        {
            foreach($page in $InputObject.pages)
            {
                ConvertTo-Page `
                    -InputObject $page `
                    -SpaceKey $SpaceKey `
                    -ParentPageId $newPage.id
            }
        }

         if ($InputObject.type -eq "Page")
        {
            foreach($page in $InputObject.pages)
            {
                ConvertTo-Page `
                    -InputObject $page `
                    -SpaceKey $SpaceKey `
                    -ParentPageId $newPage.id
            }
        }
    }
}
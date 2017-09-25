[CmdletBinding()]
Param(
    [Parameter(Mandatory = $true)]
    [string]$LiteralPath
)
Begin
{
    . (Join-Path $PSScriptRoot .\New-Space.ps1)
    . (Join-Path $PSScriptRoot .\ConvertTo-Page.ps1)
}
Process
{
    $credential = `
        Get-Credential `
            -UserName "pavels.ahmadulins" `
            -Message "Please provide password for Confluence user"
            
    $notebookNum = 46
    $notebooks = Get-Content -LiteralPath $LiteralPath | ConvertFrom-Json

    foreach($notebook in $notebooks)
    {
        Write-Verbose "Creating '$($notebook.title)' notebook..."

        $notebookKey = "ONENOTE$notebookNum"
        $notebookNum++

        New-Space `
            -Key $notebookKey `
            -Title ($notebook.title + "19") `
            -Credential $credential `
            | Out-Null

        $Global:pageNameTitle = @{}

        foreach($section in $notebook.sections)
        {
            ConvertTo-Page `
                -InputObject $section `
                -SpaceKey $notebookKey
        }
    }
}
End
{

}
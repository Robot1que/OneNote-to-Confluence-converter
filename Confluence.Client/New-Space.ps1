function New-Space
{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [string]$Key,
        [Parameter(Mandatory = $true)]
        [string]$Title,
        [Parameter(Mandatory = $true)]
        [pscredential]$Credential
    )
    Begin
    {
        . (Join-Path $PSScriptRoot .\Post-WebRequest.ps1)
    }
    Process
    {
        $Body = @{
            key = $Key;
            name = $Title;
        }

        Post-WebRequest `
            -Credential $Credential `
            -Uri "http://10.1.2.143:8090/rest/api/space/" `
            -Body ($Body | ConvertTo-Json)
    }
    End
    {
        
    }
}
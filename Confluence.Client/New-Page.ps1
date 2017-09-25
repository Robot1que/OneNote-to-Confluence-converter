function New-Page
{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [string]$Title,
        [Parameter(Mandatory = $true)]
        [string]$SpaceKey,
        [Parameter(Mandatory = $false)]
        [int]$ParentPageId,
        [Parameter(Mandatory = $true)]
        [string]$Content,
        [Parameter(Mandatory = $true)]
        [pscredential]$Credential
    )
    Begin
    {
        . (Join-Path $PSScriptRoot .\Post-WebRequest.ps1)
    }
    Process
    {
        $body = @{
            type = "page";
            title = $Title;
            space = @{ key = $SpaceKey };
            body = @{
                storage = @{
                    value = $Content;
                    representation = "storage"
                }
            }
        }

        if ($ParentPageId -ne $null)
        {
            $body.Add("ancestors", @( @{ id = $ParentPageId } ))
        }

        Post-WebRequest `
            -Credential $Credential `
            -Uri "http://10.1.2.143:8090/rest/api/content/" `
            -Body ($body | ConvertTo-Json)
    }
}


#  {"type":"page","title":"new page", 
#  "space":{"key":"TST"},"body":{"storage":{"value":
# "<p>This is a new page</p>","representation":"storage"}}}
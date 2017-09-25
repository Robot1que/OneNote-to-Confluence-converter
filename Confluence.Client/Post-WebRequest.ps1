function Post-WebRequest
{
    [CmdletBinding()]
    Param(
        [Parameter(Mandatory = $true)]
        [pscredential]$Credential,
        [Parameter(Mandatory = $true)]
        [string]$Uri,
        [Parameter(Mandatory = $true)]
        [string]$Body
    )
    Process
    {
        $pair = "$($Credential.UserName):$($Credential.GetNetworkCredential().password)"
        $encodedCreds = [System.Convert]::ToBase64String([System.Text.Encoding]::ASCII.GetBytes($pair))

        $Headers = @{}
        $Headers.Add("Authorization", "Basic $encodedCreds")
        $Headers.Add("Content-Type", "application/json")

        $response =Invoke-WebRequest -Method "POST" -Uri $Uri -Headers $Headers -Body $Body

        if ($response -ne $null -and $response.StatusCode -eq 200)
        {
            $response.Content | ConvertFrom-Json
        }
    }
}
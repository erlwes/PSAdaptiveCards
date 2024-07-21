function Invoke-PostJsonToWebhook {
    param (
        [Parameter(Mandatory = $true)]
        [string]$WebhookURI,

        [Parameter(ValueFromPipeline = $true, Mandatory = $true)]
        [string]$Json
    )
    begin {
        $parameters = @{
            "URI"         = $WebhookURI
            "Method"      = 'POST'
            "Body"        = $Json
            "ContentType" = 'application/json; charset=UTF-8'
            "ErrorAction" = 'Stop'
        }
    }
    process {
        try {            
            Invoke-RestMethod @parameters            
        } catch {
            Write-Error "Failed to send request: $($_.Exception.Message)"
        }
    }
}
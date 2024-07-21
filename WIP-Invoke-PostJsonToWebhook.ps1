# Webhook URI will be generated when you add a workflow to your teams-channel and select the template "post-to-a-channel-when-a-webhook-request-is-received"
# Read more here: https://devblogs.microsoft.com/microsoft365dev/retirement-of-office-365-connectors-within-microsoft-teams/
# Example URI: 'https://prod-140.westeurope.logic.azure.com:443/workflows/[REDACTED]/triggers/manual/paths/invoke?api-version=2016-06-01&sp=%2Ftriggers%2Fmanual%2Frun&sv=1.0&sig=[REDACTED]'

# THIS FUNTION NOT TESTED WITH PIPELINES AND OUTPUT FROM New-AdaptiveCard YET! SHOULD BE STRAIGHT FORWARD.

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

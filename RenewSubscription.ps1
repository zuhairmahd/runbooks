<#
.SYNOPSIS
    Azure Function to automatically renew Microsoft Graph change notification subscriptions.

.DESCRIPTION
    This timer-triggered function runs every 12 hours to check and renew Graph API subscriptions
    before they expire. It reads subscription information from Azure App Configuration or
    environment variables and renews subscriptions that are close to expiring.

.NOTES
    - Runs on a timer trigger (every 12 hours by default)
    - Renews subscriptions that expire within 24 hours
    - Requires Microsoft.Graph.ChangeNotifications module
    - Requires appropriate Graph API permissions
#>
[CmdletBinding()]
param()

Write-Host "================================================"
Write-Host "Microsoft Graph Subscription Renewal Function"
Write-Host "================================================"
Write-Host "Execution time: $((Get-Date).ToString('o'))"

# Get client ID from environment variable or App Configuration
$managedIdentityClientId = $env:change_group_function_identity_client_id
$subscriptionRenewalPeriodHours = if ($env:SUBSCRIPTION_RENEWAL_PERIOD_HOURS)
{
    [int]$env:SUBSCRIPTION_RENEWAL_PERIOD_HOURS
}
else
{
    24
}

# Try to find subscription by querying all subscriptions for this resource
try
{
    Connect-MgGraph -Identity -ClientId $managedIdentityClientId -NoWelcome
    Write-Host "Connected to Microsoft Graph"

    $allSubscriptions = Get-MgSubscription -All
    Write-Host "Got $($allSubscriptions.Count) total subscriptions  "

    # Filter for subscriptions that use EventGrid and match our resource
    $relevantSubscriptions = $allSubscriptions | Where-Object {
        $_.NotificationUrl -like "*EventGrid*" -and
        $_.Resource -eq "groups"
    }
    Write-Host "Filtering for subscriptions with resource 'groups' and EventGrid notification URL"
    if ($relevantSubscriptions.Count -eq 0)
    {
        Write-Host "No subscriptions found - creating new one..." -ForegroundColor Yellow
        $subscriptionId = $env:AZURE_SUBSCRIPTION_ID ?? "8a89e116-824d-4eeb-8ef4-16dcc1f0959b"
        $resourceGroup = $env:RESOURCE_GROUP_NAME ?? "groupchangefunction"
        $partnerTopic = $env:PARTNER_TOPIC_NAME ?? "default"
        $location = $env:AZURE_REGION ?? "centralus"

        $newExpiration = (Get-Date).AddMinutes(4230)
        $clientState = [Guid]::NewGuid().ToString()

        $createParams = @{
            changeType               = "updated,deleted,created"
            notificationUrl          = "EventGrid:?azuresubscriptionid=$subscriptionId&resourcegroup=$resourceGroup&partnertopic=$partnerTopic&location=$location"
            lifecycleNotificationUrl = "EventGrid:?azuresubscriptionid=$subscriptionId&resourcegroup=$resourceGroup&partnertopic=$partnerTopic&location=$location"
            resource                 = "groups"
            expirationDateTime       = $newExpiration
            clientState              = $clientState
        }

        try
        {
            $newSubscription = New-MgSubscription -BodyParameter $createParams
            Write-Host "✅ Created new subscription: $($newSubscription.Id)" -ForegroundColor Green
            Write-Host "   Expires: $($newSubscription.ExpirationDateTime)" -ForegroundColor Green
            Write-Host "⚠️  IMPORTANT: New subscription ID generated" -ForegroundColor Yellow
            Write-Host "   To optimize future runs, set GRAPH_SUBSCRIPTION_ID in Function App settings:" -ForegroundColor Yellow
            Write-Host "   Value: $($newSubscription.Id)" -ForegroundColor White

            # Continue processing with this new subscription
            $relevantSubscriptions = @($newSubscription)
        }
        catch
        {
            Write-Error "Failed to create subscription: $($_.Exception.Message)"
            throw
        }
    }

    Write-Host "Found $($relevantSubscriptions.Count) relevant subscription(s)"
    foreach ($sub in $relevantSubscriptions)
    {
        $graphSubscriptionId = $sub.Id
        Write-Host "`nProcessing subscription: $graphSubscriptionId"

        $expirationDateTime = [DateTime]::Parse($sub.ExpirationDateTime)
        $hoursUntilExpiration = ($expirationDateTime - (Get-Date)).TotalHours

        Write-Host "  Resource: $($sub.Resource)"
        Write-Host "  Expiration: $expirationDateTime"
        Write-Host "  Hours until expiration: $([Math]::Round($hoursUntilExpiration, 2))"

        # Renew if expiring within subscription renewal period hours
        if ($hoursUntilExpiration -lt $subscriptionRenewalPeriodHours)
        {
            Write-Host "  ⚠️  Subscription expires soon! Renewing..." -ForegroundColor Yellow
            # Set new expiration to maximum (4230 minutes)
            $newExpiration = (Get-Date).AddMinutes(4230)

            $updateParams = @{
                ExpirationDateTime = $newExpiration
            }

            try
            {
                $updated = Update-MgSubscription -SubscriptionId $graphSubscriptionId -BodyParameter $updateParams
                Write-Host "  ✅ Subscription renewed successfully!" -ForegroundColor Green
                Write-Host "  New expiration: $($updated.ExpirationDateTime)" -ForegroundColor Green
            }
            catch
            {
                $errorMsg = $_.Exception.Message
                Write-Error "  ❌ Failed to renew subscription: $errorMsg"
                # If subscription is expired, invalid, or doesn't belong to this app
                if ($errorMsg -like "*ResourceNotFound*" -or $errorMsg -like "*expired*" -or $errorMsg -like "*does not belong to application*")
                {
                    Write-Warning "  The subscription may have expired or was created by a different application."
                    Write-Warning "  Run create-api-subscription-topic.ps1 to create a new subscription."
                    Write-Warning "  For Function App: Subscription must be created using the managed identity."
                }
            }
        }
        else
        {
            Write-Host "  ✅ Subscription is still valid (expires in $([Math]::Round($hoursUntilExpiration, 1)) hours)" -ForegroundColor Green
        }
    }
}
catch
{
    Write-Error "Failed to connect to Microsoft Graph or query subscriptions: $($_.Exception.Message)"
    throw
}

Write-Host "`n================================================"
Write-Host "Renewal check completed at $((Get-Date).ToString('o'))"
Write-Host "================================================"


Disconnect-MgGraph
Write-Host "`nDisconnected from Microsoft Graph"

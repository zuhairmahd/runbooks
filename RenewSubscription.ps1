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

Write-Output "================================================"
Write-Output "Microsoft Graph Subscription Renewal Function"
Write-Output "================================================"
Write-Output "Execution time: $((Get-Date).ToString('o'))"

#region Get automation variables
$managedIdentityClientId = Get-AutomationVariable -Name 'change_group_function_identity_client_id'
$subscriptionRenewalPeriodHours = if (Get-AutomationVariable -Name 'SUBSCRIPTION_RENEWAL_PERIOD_HOURS' -ErrorAction SilentlyContinue)
{
    [int](Get-AutomationVariable -Name 'SUBSCRIPTION_RENEWAL_PERIOD_HOURS')
}
else
{
    24
}

$subscriptionId = Get-AutomationVariable -Name 'AZURE_SUBSCRIPTION_ID' -ErrorAction SilentlyContinue
if (-not $subscriptionId)
{
    Write-Output "AZURE_SUBSCRIPTION_ID variable not set."
    $subscriptionId = $null
}

$resourceGroup = Get-AutomationVariable -Name 'AZURE_RESOURCE_GROUP' -ErrorAction SilentlyContinue
if (-not $resourceGroup)
{
    $resourceGroup = 'groupchangefunction'
    Write-Output "AZURE_RESOURCE_GROUP variable not set. Using default: $resourceGroup"
}

$partnerTopic = Get-AutomationVariable -Name 'EVENT_GRID_PARTNER_TOPIC' -ErrorAction SilentlyContinue
if (-not $partnerTopic)
{
    $partnerTopic = 'default'
    Write-Output "EVENT_GRID_PARTNER_TOPIC variable not set. Using default: $partnerTopic"
}

$partnerTopicId = Get-AutomationVariable -Name 'EVENT_GRID_PARTNER_TOPIC_ID' -ErrorAction SilentlyContinue
if (-not $partnerTopicId)
{
    $partnerTopicId = $null
    Write-Output "EVENT_GRID_PARTNER_TOPIC_ID variable not set. Using default: $partnerTopicId"
}

$location = Get-AutomationVariable -Name 'AZURE_LOCATION' -ErrorAction SilentlyContinue
if (-not $location)
{
    $location = 'centralus'
    Write-Host "AZURE_LOCATION variable not set. Using default: $location"
}
#endregion Get automation variables

try
{
    Connect-MgGraph -Identity -ClientId $managedIdentityClientId -NoWelcome
    Write-Output "Connected to Microsoft Graph"
    if (-not $partnerTopicId )
    {
        # Try to find partner topic subscription by querying all subscriptions for this resource
        $allPartnerTopicSubscriptions = Get-MgSubscription -All
        Write-Output "Got response: $($relevantSubscriptions |Out-String)"

        # Filter for a subscription that uses EventGrid and match our resource
        Write-Output "Filtering for subscriptions with resource 'groups' and EventGrid notification URL"
        $relevantSubscriptions = $allPartnerTopic       Subscriptions | Where-Object {
            $_.NotificationUrl -like "*EventGrid*" -and
            $_.Resource -eq "groups"
        } | Select-Object -First 1
        if ($null -eq $relevantSubscriptions)
        {
            Write-Output "No subscriptions found - creating new one..."

            # If subscription ID is not configured, try to get it from the current Azure context
            if ([string]::IsNullOrWhiteSpace($subscriptionId))
            {
                Write-Output "Subscription ID not configured in automation variables. Attempting to retrieve from current Azure context..."
                try
                {
                    $context = Get-MgContext
                    if ($context)
                    {
                        $subscriptionId = $context.TenantId
                        Write-Output "Retrieved Tenant ID from Microsoft Graph context: $subscriptionId"
                    }
                }
                catch
                {
                    Write-Warning "Could not retrieve subscription ID from context: $($_.Exception.Message)"
                }
            }

            # Validate subscription ID format (must be a valid GUID)
            if ([string]::IsNullOrWhiteSpace($subscriptionId))
            {
                Write-Error "AZURE_SUBSCRIPTION_ID automation variable is not set and could not be auto-detected from context."
                Write-Error "Please set the AZURE_SUBSCRIPTION_ID variable in your Automation Account or ensure the Azure context is properly established."
                throw "Missing required subscription ID"
            }

            $newExpiration = (Get-Date).AddMinutes(4230)
            $clientState = [Guid]::NewGuid().ToString()
            Write-Output "Creating subscription with:"
            Write-Output "  Subscription ID: $subscriptionId"
            Write-Output "  Resource Group: $resourceGroup"
            Write-Output "  Partner Topic: $partnerTopic"
            Write-Output "  Location: $location"
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
                Write-Output "Created new subscription: $($newSubscription.Id)"
                Write-Output "Expires: $($newSubscription.ExpirationDateTime)"
                Write-Output "IMPORTANT: New subscription ID generated"
                Write-Output "To optimize future runs, set AZURE_TOPIC_SUBSCRIPTION_ID in Automation Account variables:"
                Write-Output "Value: $($newSubscription.Id)"
                # Continue processing with this new subscription
                $relevantSubscriptions = @($newSubscription)
            }
            catch
            {
                Write-Error "Failed to create subscription: $($_.Exception.Message)"
                throw
            }
        }
        else
        {
            $graphSubscriptionId = $relevantSubscriptions.Id
            $expirationDateTime = [DateTime]$relevantSubscriptions.ExpirationDateTime
        }
    }
    else
    {
        Write-Output "Getting subscription by ID: $partnerTopicId"
        $sub = Get-MgSubscription -SubscriptionId $partnerTopicId
        if ($sub)
        {
            $graphSubscriptionId = $sub.Id
            $expirationDateTime = [DateTime]$sub.ExpirationDateTime
            Write-Output "`nProcessing subscription: $graphSubscriptionId"
        }
    }

    if ($graphSubscriptionId -and $expirationDateTime)
    {
        $hoursUntilExpiration = ($expirationDateTime - (Get-Date)).TotalHours
        Write-Output "  Resource: $($sub.Resource)"
        Write-Output "  Expiration: $expirationDateTime"
        Write-Output "  Hours until expiration: $([Math]::Round($hoursUntilExpiration, 2))"

        # Renew if expiring within subscription renewal period hours
        if ($hoursUntilExpiration -lt $subscriptionRenewalPeriodHours)
        {
            Write-Output "  ⚠️  Subscription expires soon! Renewing..."            # Set new expiration to maximum (4230 minutes)
            $newExpiration = (Get-Date).AddMinutes(4230)

            $updateParams = @{
                ExpirationDateTime = $newExpiration
            }

            try
            {
                $updated = Update-MgSubscription -SubscriptionId $graphSubscriptionId -BodyParameter $updateParams
                Write-Output "  ✅ Subscription renewed successfully!" Write-Output "  New expiration: $($updated.ExpirationDateTime)"
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
            Write-Output "  ✅ Subscription is still valid (expires in $([Math]::Round($hoursUntilExpiration, 1)) hours)"
        }
    }
}
catch
{
    Write-Error "Failed to connect to Microsoft Graph or query subscriptions: $($_.Exception.Message)"
    throw
}
finally
{
    Write-Output "`n================================================"
    Write-Output "Renewal check completed at $((Get-Date).ToString('o'))"
    Write-Output "================================================"
    Disconnect-MgGraph
    Write-Output "`nDisconnected from Microsoft Graph"
}
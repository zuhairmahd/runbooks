<#
.SYNOPSIS
    Renews Microsoft Graph group-change subscriptions that deliver notifications through Azure Event Grid.

.DESCRIPTION
    Ensures the Graph subscription that posts group change notifications to Event Grid stays active. The
    runbook connects to Microsoft Graph with a managed identity, locates the current subscription (by ID or
    discovery), and renews it when the expiration window is within the configured threshold. If no matching
    subscription exists, it creates a new one with the Event Grid endpoint parameters derived from automation
    variables or environment variables.

.PARAMETER WhatIf
    Shows the actions that would be taken (create or renew) without performing them.

.PARAMETER Disconnect
    Disconnects the Microsoft Graph session at the end of the run. Automatically enabled in Azure Automation.

.NOTES
    - Requires: Microsoft Graph PowerShell module and managed identity authentication
    - Automation / environment variables:
      * change_group_function_identity_client_id (required in Azure Automation): Managed identity client ID
      * SUBSCRIPTION_RENEWAL_PERIOD_HOURS (optional): Hours before expiration to trigger renewal (default: 24)
      * AZURE_SUBSCRIPTION_ID (optional): Azure subscription ID for Event Grid
      * AZURE_RESOURCE_GROUP (optional): Resource group name (default: groupchangefunction)
      * EVENT_GRID_PARTNER_TOPIC (optional): Partner topic name (default: default)
      * EVENT_GRID_PARTNER_TOPIC_ID (optional): Subscription ID for direct lookup
      * AZURE_LOCATION (optional): Azure region (default: centralus)

.EXAMPLE
    .\RenewSubscription.ps1 -Verbose
    Runs locally, uses environment variables, and prints verbose discovery/renewal details.

.EXAMPLE
    .\RenewSubscription.ps1 -WhatIf
    Shows what renewal or creation would occur without changing the Graph subscription.

.FUNCTIONALITY
    - Connects to Microsoft Graph using managed identity
    - Discovers an existing group-change subscription or creates one when missing
    - Renews subscriptions that are within the configured expiration window
    - Emits clear guidance when the subscription belongs to a different app or is expired
    - Provides detailed output for monitoring and troubleshooting
#>
[CmdletBinding()]
param(
    [switch]$WhatIf,
    [switch]$Disconnect
)

#region Helper functions
function Test-IsAzureAutomation()
{
    [CmdletBinding()]
    param()

    # Detect if running in Azure Automation or local terminal
    $isAzureAutomation = $false
    $functionName = $MyInvocation.MyCommand.Name
    # Check for Azure Automation specific environment variables and context
    if ($env:AUTOMATION_ASSET_ACCOUNTID -or
        $env:AZUREPS_HOST_ENVIRONMENT -eq 'AzureAutomation' -or
        (Get-Command Get-AutomationVariable -ErrorAction SilentlyContinue))
    {
        Write-Verbose "[$functionName] Detected Azure Automation environment variables or cmdlets."
        $isAzureAutomation = $true
    }
    else
    {
        Write-Verbose "[$functionName] No Azure Automation environment variables or cmdlets detected."
    }
    return $isAzureAutomation
}

function Get-VariableValue()
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$VariableName,
        [Parameter(Mandatory = $false)]
        [object]$DefaultValue,
        [Parameter(Mandatory = $false)]
        [bool]$IsAzureAutomation
    )

    $value = $null
    $functionName = $MyInvocation.MyCommand.Name
    Write-Verbose "[$functionName] Retrieving variable '$VariableName' (IsAzureAutomation: $IsAzureAutomation)"
    if ($IsAzureAutomation)
    {
        $value = Get-AutomationVariable -Name $VariableName -ErrorAction SilentlyContinue
        Write-Verbose "[$functionName] Retrieved value from Automation Variable: $value"
    }
    else
    {
        $value = [System.Environment]::GetEnvironmentVariable($VariableName)
        Write-Verbose "[$functionName] Retrieved value from local Environment Variable: $value"
    }

    if (-not $value -and $null -ne $DefaultValue)
    {
        Write-Verbose "[$functionName] Variable '$VariableName' not set. Using default value: $DefaultValue"
        return $DefaultValue
    }

    return $value
}

function Connect-ToAzure()
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]$ClientId,
        [Parameter(Mandatory = $true)]
        [bool]$IsAzureAutomation,
        [switch]$useSystemAssignedIdentity
    )

    $functionName = $MyInvocation.MyCommand.Name
    $context = Get-MgContext
    if ($context)
    {
        Write-Verbose "[$functionName] Already connected to Microsoft Graph with context: $($context.Account)"
        #output basic context
        Write-Output "[$functionName] Current Microsoft Graph context:"
        Write-Output "  Account: $($context.Account)"
        Write-Output "  TenantId: $($context.TenantId)"
        Write-Output "  Environment: $($context.Environment)"
        Write-Output "Auth type: $($context.AuthType)"
        Write-Output "  Scopes: $($context.Scopes -join ', ')"
    }
    else
    {
        Write-Verbose "[$functionName] Connecting to Microsoft Graph (IsAzureAutomation: $IsAzureAutomation)"
        Write-Verbose "[$functionName] ClientId: $(if ($ClientId) { 'Provided' } else { 'Not Provided' })"
        if ($IsAzureAutomation)
        {
            if ($useSystemAssignedIdentity)
            {
                Write-Verbose "[$functionName] Using system-assigned managed identity to connect to Microsoft Graph."
                try
                {
                    Connect-MgGraph -Identity -NoWelcome -ErrorAction Stop
                    Write-Verbose "[$functionName] Connected to Microsoft Graph using system-assigned managed identity."
                }
                catch
                {
                    Write-Error "[$functionName] Failed to connect to Microsoft Graph using system-assigned managed identity: $($_.Exception.Message)"
                    throw
                }
            }
            else
            {
                Write-Verbose "[$functionName] Using user-assigned managed identity with ClientId: $ClientId to connect to Microsoft Graph."
                try
                {
                    Connect-MgGraph -Identity -ClientId $ClientId -NoWelcome -ErrorAction Stop
                    Write-Output "[$functionName] Connected to Microsoft Graph using managed identity."
                }
                catch
                {
                    Write-Error "[$functionName] Failed to connect to Microsoft Graph using managed identity: $($_.Exception.Message)"
                    throw
                }
            }
        }
        else
        {
            # Local terminal: use default Azure context
            Connect-MgGraph -NoWelcome -ErrorAction Stop
            Write-Verbose "[$functionName] Connected to Microsoft Graph using default Azure context."
        }
    }
    return Get-MgContext
}
#endregion Helper functions

#region Get automation variables
# Detect execution environment
$isAzureAutomation = Test-IsAzureAutomation
$managedIdentityClientId = Get-VariableValue -VariableName 'change_group_function_identity_client_id' -IsAzureAutomation $isAzureAutomation
$subscriptionRenewalPeriodHours = [int](Get-VariableValue -VariableName 'SUBSCRIPTION_RENEWAL_PERIOD_HOURS' -DefaultValue 24 -IsAzureAutomation $isAzureAutomation)
$subscriptionId = Get-VariableValue -VariableName 'AZURE_SUBSCRIPTION_ID' -DefaultValue "8a89e116-824d-4eeb-8ef4-16dcc1f0959b" -IsAzureAutomation $isAzureAutomation
$resourceGroup = Get-VariableValue -VariableName 'AZURE_RESOURCE_GROUP' -DefaultValue 'groupchangefunction' -IsAzureAutomation $isAzureAutomation
$partnerTopic = Get-VariableValue -VariableName 'EVENT_GRID_PARTNER_TOPIC' -DefaultValue 'default' -IsAzureAutomation $isAzureAutomation
$partnerTopicId = Get-VariableValue -VariableName 'EVENT_GRID_PARTNER_TOPIC_ID' -IsAzureAutomation $isAzureAutomation
$location = Get-VariableValue -VariableName 'AZURE_LOCATION' -DefaultValue 'centralus' -IsAzureAutomation $isAzureAutomation
#validate the variables for good measure
if (-not $subscriptionId)
{
    Write-Output "AZURE_SUBSCRIPTION_ID variable not set. Will attempt to autodetect."
    throw "Missing required variable: AZURE_SUBSCRIPTION_ID"
}
if (-not $managedIdentityClientId -and $isAzureAutomation)
{
    Write-Error "change_group_function_identity_client_id automation variable is required in Azure Automation."
    throw "Missing required variable: change_group_function_identity_client_id"
}
#endregion Get automation variables

Write-Output "================================================"
Write-Output "Microsoft Graph Subscription Renewal Function"
Write-Output "================================================"
Write-Output "Execution time: $((Get-Date).ToString('o'))"
Write-Output "Execution Environment: $(if ($isAzureAutomation) { 'Azure Automation' } else { 'Local Terminal' })"

try
{
    $connectionContext = Connect-ToAzure -ClientId $managedIdentityClientId -IsAzureAutomation $isAzureAutomation
    if ($connectionContext)
    {
        Write-Output "Connected to Microsoft Graph"
    }
    else
    {
        throw "Failed to connect to Microsoft Graph"
    }

    if (-not $partnerTopicId )
    {
        # Try to find partner topic subscription by querying all subscriptions for this resource
        $allPartnerTopicSubscriptions = Get-MgSubscription -All
        Write-Verbose "Got response: $($allPartnerTopicSubscriptions |Out-String)"

        # Filter for a subscription that uses EventGrid and match our resource
        Write-Output "Filtering for subscriptions with resource 'groups' and EventGrid notification URL"
        $relevantSubscriptions = $allPartnerTopicSubscriptions | Where-Object {
            $_.NotificationUrl -like "*EventGrid*" -and
            $_.Resource -eq "groups"
        } | Select-Object -First 1
        if ($null -eq $relevantSubscriptions)
        {
            Write-Output "No subscriptions found - creating new one..."

            # If subscription ID is not configured, try to get it from the current Azure context

            $newExpiration = (Get-Date).AddMinutes(4230)
            $clientState = [Guid]::NewGuid().ToString()
            Write-Output "Creating subscription with:"
            Write-Output "  Subscription ID: $subscriptionId"
            Write-Output "  Resource Group: $resourceGroup"
            Write-Output "  Partner Topic: $partnerTopic"
            Write-Output "  Location: $location"
            if ($whatIf)
            {
                Write-Output "  WHATIF: New subscription would be created with expiration $newExpiration"
            }
            else
            {
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
        }
        else
        {
            $graphSubscriptionObject = $relevantSubscriptions
        }
    }
    else
    {
        Write-Output "Getting subscription by ID: $partnerTopicId"
        $sub = Get-MgSubscription -SubscriptionId $partnerTopicId
        if ($sub)
        {
            $graphSubscriptionObject = $sub
        }
    }

    if ($graphSubscriptionObject)
    {
        $hoursUntilExpiration = ($graphSubscriptionObject.ExpirationDateTime - (Get-Date)).TotalHours
        Write-Output "Found Resource Graph for Endpoint: /$($graphSubscriptionObject.Resource)"
        Write-Output "  Expiration Date: $($graphSubscriptionObject.ExpirationDateTime)"
        Write-Output "  Hours until expiration: $([Math]::Round($hoursUntilExpiration, 2))"
        # Renew if expiring within subscription renewal period hours
        if ($hoursUntilExpiration -lt $subscriptionRenewalPeriodHours)
        {
            Write-Output "  Subscription expires soon! Renewing..." # Set new expiration to maximum (4230 minutes)
            $newExpiration = (Get-Date).AddMinutes(4230)
            if ($whatIf)
            {
                Write-Output "  WHATIF: Subscription would be renewed with new expiration $newExpiration"
            }
            else
            {
                $updateParams = @{
                    ExpirationDateTime = $newExpiration
                }
                try
                {
                    $updated = Update-MgSubscription -SubscriptionId $graphSubscriptionObject.Id -BodyParameter $updateParams
                    Write-Output "   Subscription renewed successfully!"
                    Write-Output "  New expiration: $($updated.ExpirationDateTime)"
                }
                catch
                {
                    $errorMsg = $_.Exception.Message
                    Write-Error "   Failed to renew subscription: $errorMsg"
                    # If subscription is expired, invalid, or doesn't belong to this app
                    if ($errorMsg -like "*ResourceNotFound*" -or $errorMsg -like "*expired*" -or $errorMsg -like "*does not belong to application*")
                    {
                        Write-Warning "  The subscription may have expired or was created by a different application."
                        Write-Warning "  For Function App: Subscription must be created using the managed identity."
                    }
                }
            }
        }
        else
        {
            Write-Output "   Subscription is still valid (expires in $([Math]::Round($hoursUntilExpiration, 1)) hours)"
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
    if ($isAzureAutomation -or $disconnect)
    {
        Disconnect-MgGraph
        Write-Output "`nDisconnected from Microsoft Graph"
    }
}


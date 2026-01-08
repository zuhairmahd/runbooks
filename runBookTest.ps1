function Test-IsAzureAutomation()
{
    [CmdletBinding()                        ]
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
        Write-Verbose "[$functionName]                                  No Azure Automation environment variables or cmdlets detected."
    }
    return $isAzureAutomation
}
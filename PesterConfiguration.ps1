<#
.SYNOPSIS
    Central Pester configuration for Autopilot test suite
.DESCRIPTION
    Defines standard Pester configuration used across all test executions
    Compatible with PowerShell 5.1 and Pester 5.x
#>

function Get-AutopilotPesterConfiguration
{
    [CmdletBinding()]
    param(
        [ValidateSet('Unit', 'Integration', 'Comprehensive', 'All')]
        [string]$TestType = 'All',
        [ValidateSet('None', 'Minimal', 'Normal', 'Detailed')]
        [string]$OutputVerbosity = 'Normal',
        [switch]$EnableCodeCoverage,
        [switch]$CI,
        [switch]$Exclude
    )
    
    $config = New-PesterConfiguration
    
    # Output configuration
    $config.Output.Verbosity = $OutputVerbosity
    
    # Test discovery - handle exclusion mode
    if ($Exclude -and $TestType -ne 'All')
    {
        # Exclusion mode: run everything EXCEPT the specified type
        switch ($TestType)
        {
            'Unit'
            {
                $config.Run.Path = @('.\tests\Integration', '.\tests\Comprehensive')
                $config.Filter.Tag = @('Integration', 'Comprehensive')
            }
            'Integration'
            {
                $config.Run.Path = @('.\tests\Unit', '.\tests\Comprehensive')
                $config.Filter.Tag = @('Unit', 'Comprehensive')
            }
            'Comprehensive'
            {
                $config.Run.Path = @('.\tests\Unit', '.\tests\Integration')
                $config.Filter.Tag = @('Unit', 'Integration')
            }
        }
        # Note: -TestType All with -Exclude is not meaningful, so it's ignored
        # and behaves the same as -TestType All without -Exclude
    }
    elseif ($TestType -ne 'All')
    {
        # Normal inclusion mode
        switch ($TestType)
        {
            'Unit'
            {
                $config.Run.Path = '.\tests\Unit'
                $config.Filter.Tag = 'Unit'
            }
            'Integration'
            {
                $config.Run.Path = '.\tests\Integration'
                $config.Filter.Tag = 'Integration'
            }
            'Comprehensive'
            {
                $config.Run.Path = '.\tests\Comprehensive'
                $config.Filter.Tag = 'Comprehensive'
            }
        }
    }
    else
    {
        # Run all tests
        $config.Run.Path = '.\tests'
        # Exclude template and helper files from test execution
        $config.Run.ExcludePath = @('*Template.Tests.ps1', '*Helpers*')
    }
    
    # Test execution
    $config.Run.Exit = $false  # Don't exit PowerShell after tests
    $config.Run.PassThru = $true  # Return test results object
    
    # TestDrive configuration - ensure cleanup
    $config.TestDrive.Enabled = $true
    
    # Code coverage
    if ($EnableCodeCoverage)
    {
        $config.CodeCoverage.Enabled = $true
        $config.CodeCoverage.Path = '.\functions\**\*.ps1'
        $config.CodeCoverage.OutputPath = '.\coverage.xml'
        $config.CodeCoverage.OutputFormat = 'JaCoCo'
    }
    
    # CI/CD integration
    if ($CI)
    {
        $config.TestResult.Enabled = $true
        $config.TestResult.OutputFormat = 'NUnitXml'
        $config.TestResult.OutputPath = '.\TestResults.xml'
        $config.Output.Verbosity = 'Normal'
    }
    
    return $config
}

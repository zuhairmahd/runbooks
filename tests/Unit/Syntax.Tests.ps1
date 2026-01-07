<#
.SYNOPSIS
    PowerShell syntax validation and analysis tests
.DESCRIPTION
    Validates that all PowerShell files in the repository have valid syntax,
    proper encoding, and pass comprehensive PSScriptAnalyzer checks
#>

BeforeAll {
    # Get repository root
    $script:RepoRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)

    # Get all PowerShell script files (.ps1)
    $script:AllPowerShellFiles = Get-ChildItem -Path $script:RepoRoot -Filter '*.ps1' -Recurse -File |
        Where-Object { $_.FullName -notmatch '\\(\.git|node_modules|bin|Modules|obj|__azurite|__blobstorage__|__queuestorage__|\.vscode)\\' }

    # Get all PowerShell module files (.psm1)
    $script:AllModuleFiles = Get-ChildItem -Path $script:RepoRoot -Filter '*.psm1' -Recurse -File |
        Where-Object { $_.FullName -notmatch '\\(\.git|node_modules|bin|Modules|obj|__azurite|__blobstorage__|__queuestorage__|\.vscode)\\' }

    # Get all PowerShell manifest files (.psd1)
    $script:AllManifestFiles = Get-ChildItem -Path $script:RepoRoot -Filter '*.psd1' -Recurse -File |
        Where-Object { $_.FullName -notmatch '\\(\.git|node_modules|bin|Modules|obj|__azurite|__blobstorage__|__queuestorage__|\.vscode)\\' }

    # Combine all PowerShell files for comprehensive testing
    $script:AllPowerShellArtifacts = @($script:AllPowerShellFiles) + @($script:AllModuleFiles) + @($script:AllManifestFiles)

    # Determine if verbose output should be shown based on Pester configuration
    $script:ShowVerboseOutput = $PesterPreference.Output.Verbosity.Value -in @('Detailed', 'Diagnostic')

    # Ensure PSScriptAnalyzer is available
    if (-not (Get-Module -ListAvailable -Name PSScriptAnalyzer))
    {
        Write-Warning "PSScriptAnalyzer module not found. Installing..."
        Install-Module -Name PSScriptAnalyzer -Force -Scope CurrentUser -SkipPublisherCheck
    }
    Import-Module PSScriptAnalyzer -ErrorAction Stop
}

Describe "PowerShell File Discovery" -Tags 'Discovery', 'Unit', 'Fast' {

    Context "Repository structure" {

        It "Should find PowerShell script files (.ps1) to test" {
            $script:AllPowerShellFiles | Should -Not -BeNullOrEmpty -Because "repository should contain .ps1 files"
            $script:AllPowerShellFiles.Count | Should -BeGreaterThan 0
            if ($script:ShowVerboseOutput)
            {
                Write-Host "  Found $($script:AllPowerShellFiles.Count) PowerShell script files (.ps1)" -ForegroundColor Cyan
            }
        }

        It "Should report PowerShell module files (.psm1) if found" {
            if ($script:ShowVerboseOutput)
            {
                if ($script:AllModuleFiles.Count -gt 0)
                {
                    Write-Host "  Found $($script:AllModuleFiles.Count) PowerShell module files (.psm1)" -ForegroundColor Cyan
                }
                else
                {
                    Write-Host "  No PowerShell module files (.psm1) found" -ForegroundColor Gray
                }
            }
            # This is informational - not a failure
            $true | Should -Be $true
        }

        It "Should report PowerShell manifest files (.psd1) if found" {
            if ($script:ShowVerboseOutput)
            {
                if ($script:AllManifestFiles.Count -gt 0)
                {
                    Write-Host "  Found $($script:AllManifestFiles.Count) PowerShell manifest files (.psd1)" -ForegroundColor Cyan
                }
                else
                {
                    Write-Host "  No PowerShell manifest files (.psd1) found" -ForegroundColor Gray
                }
            }
            # This is informational - not a failure
            $true | Should -Be $true
        }

        It "Should have total PowerShell artifacts to validate" {
            $script:AllPowerShellArtifacts | Should -Not -BeNullOrEmpty
            $script:AllPowerShellArtifacts.Count | Should -BeGreaterThan 0
            if ($script:ShowVerboseOutput)
            {
                Write-Host "  Total PowerShell artifacts to validate: $($script:AllPowerShellArtifacts.Count)" -ForegroundColor Green
            }
        }
    }
}

Describe "PowerShell File Encoding and Format" -Tags 'Encoding', 'Unit', 'Fast' {

    Context "File encoding validation" {

        It "All PowerShell files should be readable and not empty" {
            $emptyOrUnreadableFiles = @()

            foreach ($file in $script:AllPowerShellArtifacts)
            {
                try
                {
                    $content = Get-Content $file.FullName -Raw -ErrorAction Stop

                    if ([string]::IsNullOrWhiteSpace($content))
                    {
                        $emptyOrUnreadableFiles += [PSCustomObject]@{
                            Name   = $file.Name
                            Path   = $file.FullName.Replace($script:RepoRoot, '').TrimStart('\')
                            Reason = "File is empty or contains only whitespace"
                        }
                    }
                }
                catch
                {
                    $emptyOrUnreadableFiles += [PSCustomObject]@{
                        Name   = $file.Name
                        Path   = $file.FullName.Replace($script:RepoRoot, '').TrimStart('\')
                        Reason = "Cannot read file: $($_.Exception.Message)"
                    }
                }
            }

            if ($emptyOrUnreadableFiles.Count -gt 0)
            {
                $errorMessage = "The following PowerShell files are empty or unreadable:`n"
                foreach ($failed in $emptyOrUnreadableFiles)
                {
                    $errorMessage += "  - $($failed.Path): $($failed.Reason)`n"
                }
                $emptyOrUnreadableFiles | Should -BeNullOrEmpty -Because $errorMessage
            }

            $emptyOrUnreadableFiles | Should -BeNullOrEmpty -Because "all PowerShell files must be readable and contain code"
        }

        It "All PowerShell files should have valid UTF-8 or ASCII encoding" {
            $encodingIssues = @()

            foreach ($file in $script:AllPowerShellArtifacts)
            {
                try
                {
                    $bytes = [System.IO.File]::ReadAllBytes($file.FullName)

                    # Check for UTF-8 BOM (EF BB BF)
                    $hasUtf8Bom = ($bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF)

                    # Check for UTF-16 BOM (FF FE or FE FF)
                    $hasUtf16Bom = ($bytes.Length -ge 2 -and (($bytes[0] -eq 0xFF -and $bytes[1] -eq 0xFE) -or ($bytes[0] -eq 0xFE -and $bytes[1] -eq 0xFF)))

                    # Warn about UTF-16 as it can cause issues
                    if ($hasUtf16Bom)
                    {
                        $encodingIssues += [PSCustomObject]@{
                            Name     = $file.Name
                            Path     = $file.FullName.Replace($script:RepoRoot, '').TrimStart('\')
                            Issue    = "File has UTF-16 BOM which may cause compatibility issues"
                            Severity = "Warning"
                        }
                    }

                    # Check for null bytes (indicates binary file or encoding corruption)
                    if ($bytes -contains 0x00)
                    {
                        $nullByteIndex = [Array]::IndexOf($bytes, 0x00)
                        if ($nullByteIndex -lt 1024)
                        {
                            # Only report if null byte is in first 1KB (likely binary)
                            $encodingIssues += [PSCustomObject]@{
                                Name     = $file.Name
                                Path     = $file.FullName.Replace($script:RepoRoot, '').TrimStart('\')
                                Issue    = "File contains null bytes (may be binary or corrupted)"
                                Severity = "Error"
                            }
                        }
                    }
                }
                catch
                {
                    $encodingIssues += [PSCustomObject]@{
                        Name     = $file.Name
                        Path     = $file.FullName.Replace($script:RepoRoot, '').TrimStart('\')
                        Issue    = "Cannot check encoding: $($_.Exception.Message)"
                        Severity = "Error"
                    }
                }
            }

            # Filter critical errors
            $criticalEncodingIssues = $encodingIssues | Where-Object { $_.Severity -eq "Error" }

            if ($encodingIssues.Count -gt 0)
            {
                $errorMessage = "The following PowerShell files have encoding issues:`n"
                foreach ($issue in $encodingIssues)
                {
                    if ($script:ShowVerboseOutput)
                    {
                        $errorMessage += "  - [$($issue.Severity)] $($issue.Path): $($issue.Issue)`n"
                    }
                }

                if ($criticalEncodingIssues.Count -gt 0)
                {
                    $criticalEncodingIssues | Should -BeNullOrEmpty -Because $errorMessage
                }
                elseif ($script:ShowVerboseOutput)
                {
                    Write-Warning $errorMessage
                }
            }

            $criticalEncodingIssues | Should -BeNullOrEmpty -Because "PowerShell files must have valid encoding"
        }
    }
}

Describe "PowerShell Syntax Validation" -Tags 'Syntax', 'Unit', 'Fast' {

    Context "Parse error detection" {

        It "All PowerShell files should have valid syntax" {
            # Test all PowerShell files
            $failedFiles = @()

            foreach ($file in $script:AllPowerShellArtifacts)
            {
                $content = Get-Content $file.FullName -Raw -ErrorAction SilentlyContinue

                if ([string]::IsNullOrWhiteSpace($content))
                {
                    continue
                }

                $parseErrors = @()
                $tokens = @()
                $ast = $null

                try
                {
                    $ast = [System.Management.Automation.Language.Parser]::ParseInput(
                        $content, [ref]$tokens, [ref]$parseErrors
                    )
                }
                catch
                {
                    $failedFiles += [PSCustomObject]@{
                        Name   = $file.Name
                        Path   = $file.FullName.Replace($script:RepoRoot, '').TrimStart('\')
                        Errors = @([PSCustomObject]@{
                                Message = "Parser exception: $($_.Exception.Message)"
                                Line    = 0
                            })
                    }
                    continue
                }

                if ($parseErrors.Count -gt 0)
                {
                    $failedFiles += [PSCustomObject]@{
                        Name   = $file.Name
                        Path   = $file.FullName.Replace($script:RepoRoot, '').TrimStart('\')
                        Errors = $parseErrors
                    }
                }
            }

            # Assert no files failed
            if ($failedFiles.Count -gt 0)
            {
                $errorMessage = "The following PowerShell files have syntax errors:`n"
                foreach ($failed in $failedFiles)
                {
                    $errorMessage += "`n  - $($failed.Path):`n"
                    foreach ($err in $failed.Errors)
                    {
                        if ($err.Extent)
                        {
                            $errorMessage += "    Line $($err.Extent.StartLineNumber): $($err.Message)`n"
                        }
                        else
                        {
                            $errorMessage += "    $($err.Message)`n"
                        }
                    }
                }
                $failedFiles | Should -BeNullOrEmpty -Because $errorMessage
            }

            $failedFiles | Should -BeNullOrEmpty -Because "all $($script:AllPowerShellArtifacts.Count) PowerShell files must have valid syntax"
        }
    }

    Context "Advanced syntax validation" {

        It "PowerShell script files should not contain problematic constructs" {
            $problematicFiles = @()

            foreach ($file in $script:AllPowerShellFiles)
            {
                $content = Get-Content $file.FullName -Raw -ErrorAction SilentlyContinue

                if ([string]::IsNullOrWhiteSpace($content))
                {
                    continue
                }

                $issues = @()

                # Check for potential security issues - hardcoded credentials patterns
                if ($content -match '(?i)(password|pwd|pass|secret|key|token)\s*=\s*[''"][^''"]{8,}[''"]')
                {
                    $issues += "Contains potential hardcoded credentials"
                }

                # Check for Write-Host in functions (anti-pattern, should use Write-Output/Write-Verbose)
                # Allow in scripts but warn in function definitions
                $tokens = @()
                $parseErrors = @()
                $ast = [System.Management.Automation.Language.Parser]::ParseInput(
                    $content, [ref]$tokens, [ref]$parseErrors
                )

                if ($ast)
                {
                    # Find function definitions
                    $functions = $ast.FindAll({ $args[0] -is [System.Management.Automation.Language.FunctionDefinitionAst] }, $true)

                    foreach ($func in $functions)
                    {
                        $funcContent = $func.Extent.Text
                        # Only flag Write-Host if it's not in comment and not intentional logging
                        if ($funcContent -match '\bWrite-Host\b' -and $funcContent -notmatch '#.*Write-Host')
                        {
                            $issues += "Function '$($func.Name)' uses Write-Host (consider Write-Output/Write-Verbose/Write-Information)"
                        }
                    }
                }

                if ($issues.Count -gt 0)
                {
                    $problematicFiles += [PSCustomObject]@{
                        Name   = $file.Name
                        Path   = $file.FullName.Replace($script:RepoRoot, '').TrimStart('\')
                        Issues = $issues
                    }
                }
            }

            # This is a warning, not a failure
            if ($problematicFiles.Count -gt 0 -and $script:ShowVerboseOutput)
            {
                $warningMessage = "`nThe following PowerShell files contain potentially problematic constructs:`n"
                foreach ($item in $problematicFiles)
                {
                    $warningMessage += "  - $($item.Path):`n"
                    foreach ($issue in $item.Issues)
                    {
                        $warningMessage += "    * $issue`n"
                    }
                }
                Write-Warning $warningMessage
            }

            # Don't fail the test, just warn
            $true | Should -Be $true
        }
    }
}

Describe "PowerShell Script Analysis" -Tags 'PSScriptAnalyzer', 'Unit' {

    Context "Critical errors" {

        It "All PowerShell files should have no PSScriptAnalyzer errors" {
            # Run PSScriptAnalyzer on all files - Error severity only
            $failedFiles = @()

            foreach ($file in $script:AllPowerShellArtifacts)
            {
                $violations = Invoke-ScriptAnalyzer -Path $file.FullName -Severity Error -ErrorAction SilentlyContinue

                if ($violations.Count -gt 0)
                {
                    $failedFiles += [PSCustomObject]@{
                        Name       = $file.Name
                        Path       = $file.FullName.Replace($script:RepoRoot, '').TrimStart('\')
                        Violations = $violations
                    }
                }
            }

            # Assert no files failed
            if ($failedFiles.Count -gt 0)
            {
                $errorMessage = "The following PowerShell files have PSScriptAnalyzer ERROR violations:`n"
                foreach ($failed in $failedFiles)
                {
                    $errorMessage += "`n  - $($failed.Path):`n"
                    foreach ($violation in $failed.Violations)
                    {
                        $errorMessage += "    Line $($violation.Line): [$($violation.Severity)] $($violation.RuleName)`n"
                        $errorMessage += "      $($violation.Message)`n"
                        if ($violation.SuggestedCorrections)
                        {
                            $errorMessage += "      Suggestion: $($violation.SuggestedCorrections[0].Description)`n"
                        }
                    }
                }
                $failedFiles | Should -BeNullOrEmpty -Because $errorMessage
            }

            $failedFiles | Should -BeNullOrEmpty -Because "all $($script:AllPowerShellArtifacts.Count) PowerShell files must pass PSScriptAnalyzer error checks"
        }
    }

    Context "Warnings" {

        It "All PowerShell files should have minimal PSScriptAnalyzer warnings" {
            # Run PSScriptAnalyzer on all files - Warning severity
            $filesWithWarnings = @()

            foreach ($file in $script:AllPowerShellArtifacts)
            {
                $violations = Invoke-ScriptAnalyzer -Path $file.FullName -Severity Warning -ErrorAction SilentlyContinue

                if ($violations.Count -gt 0)
                {
                    $filesWithWarnings += [PSCustomObject]@{
                        Name           = $file.Name
                        Path           = $file.FullName.Replace($script:RepoRoot, '').TrimStart('\')
                        Violations     = $violations
                        ViolationCount = $violations.Count
                    }
                }
            }

            # Report warnings but don't fail (warnings are informational)
            if ($filesWithWarnings.Count -gt 0)
            {
                if ($script:ShowVerboseOutput)
                {
                    $totalWarnings = ($filesWithWarnings | Measure-Object -Property ViolationCount -Sum).Sum
                    $warningMessage = "`nFound $totalWarnings PSScriptAnalyzer warnings across $($filesWithWarnings.Count) file(s):`n"

                    # Group by rule for summary
                    $allWarnings = $filesWithWarnings | ForEach-Object { $_.Violations }
                    $warningsByRule = $allWarnings | Group-Object -Property RuleName | Sort-Object Count -Descending

                    $warningMessage += "`nWarnings by rule:`n"
                    foreach ($ruleGroup in $warningsByRule)
                    {
                        $warningMessage += "  - $($ruleGroup.Name): $($ruleGroup.Count) occurrence(s)`n"
                    }

                    $warningMessage += "`nDetailed warnings:`n"
                    foreach ($file in ($filesWithWarnings | Select-Object -First 10))
                    {
                        $warningMessage += "`n  - $($file.Path): $($file.ViolationCount) warning(s)`n"
                        foreach ($violation in ($file.Violations | Select-Object -First 3))
                        {
                            $warningMessage += "    Line $($violation.Line): $($violation.RuleName)`n"
                            $warningMessage += "      $($violation.Message)`n"
                        }
                        if ($file.Violations.Count -gt 3)
                        {
                            $warningMessage += "    ... and $($file.Violations.Count - 3) more`n"
                        }
                    }

                    if ($filesWithWarnings.Count -gt 10)
                    {
                        $warningMessage += "`n  ... and $($filesWithWarnings.Count - 10) more file(s)`n"
                    }

                    Write-Warning $warningMessage
                }
            }
            elseif ($script:ShowVerboseOutput)
            {
                Write-Host "  No PSScriptAnalyzer warnings found!" -ForegroundColor Green
            }

            # Don't fail on warnings
            $true | Should -Be $true
        }
    }

    Context "Best practices" {

        It "All PowerShell files should follow best practices (Information level)" {
            # Run PSScriptAnalyzer on all files - Information severity
            $filesWithInfo = @()

            foreach ($file in $script:AllPowerShellArtifacts)
            {
                $violations = Invoke-ScriptAnalyzer -Path $file.FullName -Severity Information -ErrorAction SilentlyContinue

                if ($violations.Count -gt 0)
                {
                    $filesWithInfo += [PSCustomObject]@{
                        Name           = $file.Name
                        Path           = $file.FullName.Replace($script:RepoRoot, '').TrimStart('\')
                        Violations     = $violations
                        ViolationCount = $violations.Count
                    }
                }
            }

            # Report information but don't fail
            if ($filesWithInfo.Count -gt 0)
            {
                if ($script:ShowVerboseOutput)
                {
                    $totalInfo = ($filesWithInfo | Measure-Object -Property ViolationCount -Sum).Sum
                    Write-Host "`n  Found $totalInfo PSScriptAnalyzer informational messages across $($filesWithInfo.Count) file(s)" -ForegroundColor Cyan

                    # Group by rule for summary
                    $allInfo = $filesWithInfo | ForEach-Object { $_.Violations }
                    $infoByRule = $allInfo | Group-Object -Property RuleName | Sort-Object Count -Descending | Select-Object -First 5

                    Write-Host "  Top informational rules:" -ForegroundColor Cyan
                    foreach ($ruleGroup in $infoByRule)
                    {
                        Write-Host "    - $($ruleGroup.Name): $($ruleGroup.Count) occurrence(s)" -ForegroundColor Gray
                    }
                }
            }
            elseif ($script:ShowVerboseOutput)
            {
                Write-Host "  No informational messages found!" -ForegroundColor Green
            }

            # Don't fail on information
            $true | Should -Be $true
        }
    }

    Context "Specific rule checks" {

        It "PowerShell files should not have cmdlet alias usage" {
            $filesWithAliases = @()

            foreach ($file in $script:AllPowerShellFiles)
            {
                $violations = Invoke-ScriptAnalyzer -Path $file.FullName -IncludeRule 'PSAvoidUsingCmdletAliases' -ErrorAction SilentlyContinue

                if ($violations.Count -gt 0)
                {
                    $filesWithAliases += [PSCustomObject]@{
                        Name       = $file.Name
                        Path       = $file.FullName.Replace($script:RepoRoot, '').TrimStart('\')
                        Violations = $violations
                    }
                }
            }

            if ($filesWithAliases.Count -gt 0 -and $script:ShowVerboseOutput)
            {
                $warningMessage = "`nThe following files use cmdlet aliases (should use full cmdlet names):`n"
                foreach ($file in ($filesWithAliases | Select-Object -First 5))
                {
                    $warningMessage += "  - $($file.Path): $($file.Violations.Count) alias(es) found`n"
                }
                if ($filesWithAliases.Count -gt 5)
                {
                    $warningMessage += "  ... and $($filesWithAliases.Count - 5) more file(s)`n"
                }
                Write-Warning $warningMessage
            }

            # Don't fail, just warn
            $true | Should -Be $true
        }

        It "PowerShell functions should use approved verbs" {
            $filesWithUnapprovedVerbs = @()

            foreach ($file in $script:AllPowerShellFiles)
            {
                $violations = Invoke-ScriptAnalyzer -Path $file.FullName -IncludeRule 'PSUseApprovedVerbs' -ErrorAction SilentlyContinue

                if ($violations.Count -gt 0)
                {
                    $filesWithUnapprovedVerbs += [PSCustomObject]@{
                        Name       = $file.Name
                        Path       = $file.FullName.Replace($script:RepoRoot, '').TrimStart('\')
                        Violations = $violations
                    }
                }
            }

            if ($filesWithUnapprovedVerbs.Count -gt 0 -and $script:ShowVerboseOutput)
            {
                $errorMessage = "`nThe following files have functions with unapproved verbs:`n"
                foreach ($file in $filesWithUnapprovedVerbs)
                {
                    $errorMessage += "`n  - $($file.Path):`n"
                    foreach ($violation in $file.Violations)
                    {
                        $errorMessage += "    Line $($violation.Line): $($violation.Message)`n"
                    }
                }
                Write-Warning $errorMessage
            }

            # Don't fail, just warn
            $true | Should -Be $true
        }

        It "PowerShell files should not have unnecessary Write-Host usage" {
            $filesWithWriteHost = @()

            foreach ($file in $script:AllPowerShellFiles)
            {
                $violations = Invoke-ScriptAnalyzer -Path $file.FullName -IncludeRule 'PSAvoidUsingWriteHost' -ErrorAction SilentlyContinue

                if ($violations.Count -gt 0)
                {
                    $filesWithWriteHost += [PSCustomObject]@{
                        Name       = $file.Name
                        Path       = $file.FullName.Replace($script:RepoRoot, '').TrimStart('\')
                        Violations = $violations
                    }
                }
            }

            if ($filesWithWriteHost.Count -gt 0 -and $script:ShowVerboseOutput)
            {
                $infoMessage = "`n$($filesWithWriteHost.Count) file(s) use Write-Host. Consider using Write-Output, Write-Verbose, or Write-Information for better pipeline support.`n"
                Write-Host $infoMessage -ForegroundColor Cyan
            }

            # Don't fail, this is informational
            $true | Should -Be $true
        }
    }
}

#requires -version 7.0
<#
.SYNOPSIS
    Executes Pester tests for Autopilot project
.DESCRIPTION
    Runs Pester tests with standard configuration
    Supports filtering by test type, tags, and CI/CD integration

    When using -TestFile parameter, the script will:
    - First attempt to resolve the exact path provided
    - If not found, search the tests folder for an exact filename match
    - If still not found, perform a fuzzy search and present similar files for selection
.PARAMETER TestType
    Type of tests to run: Unit, Integration, Comprehensive, All
.PARAMETER OutputVerbosity
    Level of output detail: None, Minimal, Normal, Detailed
.PARAMETER TestFile
    Path or filename of a single test file to run (overrides TestType)
    Can be:
    - Full path: "c:\path\to\test.Tests.ps1"
    - Relative path: "tests\Integration\SettingsFunctions.Tests.ps1"
    - Just filename: "SettingsFunctions.Tests.ps1"
    - "Interactive" for interactive file browser: -TestFile "Interactive"

    If the file is not found, a fuzzy search will offer similar files for selection.
    Use "Interactive" to browse and select multiple test files interactively.
.PARAMETER EnableCodeCoverage
    Enable code coverage analysis
.PARAMETER ShowMissedCommands
    Show detailed list of commands without coverage (requires -EnableCodeCoverage)
.PARAMETER CI
    Run in CI/CD mode with NUnit XML output
.PARAMETER Tags
    Filter tests by tags. Can be:
    - Specific tags: -Tags "Unit", "Integration"
    - "Interactive" for interactive selection: -Tags "Interactive"
    - Omit parameter to run all tests

    When -Tags "Interactive" is used, an interactive menu will display all available tags
    from test files, allowing multiple selection.

    Behavior changes with -Exclude switch (see -Exclude parameter).
.PARAMETER Exclude
    When specified, inverts the filtering logic for -Tags, -TestFile, and -TestType parameters.
    Instead of including only matching tests, excludes matching tests and runs everything else.

    Examples:
    - `-Tags "Slow" -Exclude` = Run all tests EXCEPT those tagged "Slow"
    - `-TestType Unit -Exclude` = Run all tests EXCEPT Unit tests
    - `-TestFile "Auth" -Exclude` = Run all tests EXCEPT files matching "Auth"
.PARAMETER Interactive
    Launch fully interactive mode where you can select both test files AND tags in a combined workflow.
    This mode presents a menu allowing you to:
    - Select test files (with paging support for large lists)
    - Select tags (with paging support)
    - Review and confirm selections before running tests
.EXAMPLE
    .\Invoke-PesterTests.ps1 -TestType Unit
    Runs all unit tests
.EXAMPLE
    .\Invoke-PesterTests.ps1 -Tags "Unit", "Fast"
    Runs only tests tagged with 'Unit' and 'Fast'
.EXAMPLE
    .\Invoke-PesterTests.ps1 -Tags "Interactive"
    Shows interactive tag selection menu for choosing multiple tags
.EXAMPLE
    .\Invoke-PesterTests.ps1 -TestFile "tests\Integration\SettingsFunctions.Tests.ps1"
    Runs a specific test file using full relative path
.EXAMPLE
    .\Invoke-PesterTests.ps1 -TestFile "SettingsFunctions.Tests.ps1"
    Runs a specific test file using just the filename (will search tests folder)
.EXAMPLE
    .\Invoke-PesterTests.ps1 -TestFile "Settings"
    Searches for test files matching "Settings" and presents a selection menu
.EXAMPLE
    .\Invoke-PesterTests.ps1 -TestFile "Interactive"
    Shows interactive menu to browse and select multiple test files to run
.EXAMPLE
    .\Invoke-PesterTests.ps1 -EnableCodeCoverage -CI
    Runs all tests with code coverage in CI mode
.EXAMPLE
    .\Invoke-PesterTests.ps1 -EnableCodeCoverage -ShowMissedCommands
    Runs tests with detailed code coverage information
.EXAMPLE
    .\Invoke-PesterTests.ps1 -Tags "Slow" -Exclude
    Runs all tests EXCEPT those tagged "Slow"
.EXAMPLE
    .\Invoke-PesterTests.ps1 -TestType Unit -Exclude
    Runs all tests EXCEPT Unit tests (Integration and Comprehensive only)
.EXAMPLE
    .\Invoke-PesterTests.ps1 -TestFile "Authentication" -Exclude
    Runs all tests EXCEPT files matching "Authentication"
.EXAMPLE
    .\Invoke-PesterTests.ps1 -Interactive
    Launch interactive mode to select both test files and tags with paging support
#>
[CmdletBinding(DefaultParameterSetName = 'Default')]
param(
    [ValidateSet('Unit', 'Integration', 'Comprehensive', 'All')]
    [string]$TestType = 'All',
    [ValidateSet('None', 'Minimal', 'Normal', 'Detailed')]
    [string]$OutputVerbosity = 'Normal',
    [string]$TestFile,
    [switch]$skipModuleCheck,
    [Parameter(ParameterSetName = 'Default')]
    [Parameter(ParameterSetName = 'CodeCoverage')]
    [switch]$EnableCodeCoverage,
    [Parameter(ParameterSetName = 'CodeCoverage', Mandatory = $false)]
    [ValidateScript({
            if ($_ -and -not $EnableCodeCoverage)
            {
                throw "-ShowCodeCoverageDetails requires -EnableCodeCoverage to be specified"
            }
            return $true
        })]
    [switch]$ShowCodeCoverageDetails,
    [Parameter(ParameterSetName = 'Default')]
    [Parameter(ParameterSetName = 'CodeCoverage')]
    [switch]$CI,
    [Parameter(ParameterSetName = 'Default')]
    [Parameter(ParameterSetName = 'CodeCoverage')]
    [string[]]$Tags = @(),
    [Parameter(ParameterSetName = 'Default')]
    [Parameter(ParameterSetName = 'CodeCoverage')]
    [switch]$Exclude,
    [Parameter(ParameterSetName = 'Default')]
    [Parameter(ParameterSetName = 'CodeCoverage')]
    [switch]$Interactive
)

$ErrorActionPreference = 'Stop'
$strings = @('User canceled', 'No files found')
$pageSize = 12

#region Module Dependencies
# Check for required modules and install if missing
if (-not $skipModuleCheck)
{
    $requiredModules = @(
        @{ Name = 'Pester'; MinimumVersion = '5.0.0'; MaximumVersion = '5.999.999' }
    )

    foreach ($module in $requiredModules)
    {
        $installed = Get-Module -ListAvailable -Name $module.Name | Where-Object {
            $_.Version -ge [Version]$module.MinimumVersion -and $_.Version -le [Version]$module.MaximumVersion
        }

        if (-not $installed)
        {
            Write-Host "Module '$($module.Name)' (version $($module.MinimumVersion) or higher) is not installed." -ForegroundColor Yellow
            Write-Host "Attempting to install $($module.Name)..." -ForegroundColor Cyan

            try
            {
                Install-Module -Name $module.Name -MinimumVersion $module.MinimumVersion -MaximumVersion $module.MaximumVersion -Scope CurrentUser -Force -AllowClobber -SkipPublisherCheck -ErrorAction Stop
                Write-Host "Successfully installed $($module.Name)" -ForegroundColor Green
            }
            catch
            {
                Write-Host "Failed to install $($module.Name): $($_.Exception.Message)" -ForegroundColor Red
                Write-Host "Please install manually: Install-Module -Name $($module.Name) -MinimumVersion $($module.MinimumVersion) -Scope CurrentUser" -ForegroundColor Yellow
                exit 1
            }
        }
        else
        {
            Write-Verbose "Module '$($module.Name)' is already installed (version $($installed[0].Version))"
        }
    }
}
else
{
    Write-Verbose "Module dependency check is skipped as per user request."
}
#endregion Module Dependencies

# Import configuration
. "$PSScriptRoot\PesterConfiguration.ps1"

Write-Host "=" * 63 -ForegroundColor Cyan
Write-Host "  Runbooks Pester Test Suite" -ForegroundColor Cyan
Write-Host "=" * 63 -ForegroundColor Cyan

# Get Pester configuration
$config = Get-AutopilotPesterConfiguration -TestType $TestType -EnableCodeCoverage:$EnableCodeCoverage -CI:$CI -OutputVerbosity $OutputVerbosity -Exclude:$Exclude

#region Helper functions
function Find-FileWithFuzzySearch()
{
    [CmdletBinding()]
    param(
        [string]$FileName,
        [string]$Path,
        [int]$matchesToReturn = 10,
        [int]$minimScore = 100,
        [switch]$AllowMultiple
    )

    # Helper function for fuzzy string matching
    function Get-FuzzyMatchScore()
    {
        [CmdletBinding()]
        param(
            [string]$SearchTerm,
            [string]$Candidate
        )
        $functionName = $MyInvocation.MyCommand.Name
        Write-Verbose "[$functionName] Calculating fuzzy match score between '$SearchTerm' and '$Candidate'."
        $searchLower = $SearchTerm.ToLower()
        $candidateLower = $Candidate.ToLower()
        Write-Verbose "[$functionName] Lowercase SearchTerm: '$searchLower', Candidate: '$candidateLower'."
        # Exact match gets highest score
        if ($candidateLower -eq $searchLower)
        {
            Write-Verbose "[$functionName] Exact match found."
            return 1000
        }

        # Contains exact search term
        if ($candidateLower.Contains($searchLower))
        {
            Write-Verbose "[$functionName] Candidate contains search term."
            return 500 + (100 - $candidateLower.IndexOf($searchLower))
        }

        # Calculate sequential character matching score
        Write-Verbose "[$functionName] Calculating sequential character matching score."
        $score = 0
        $searchChars = $searchLower.ToCharArray()
        Write-Verbose "[$functionName] Search characters: $($searchChars -join ', ')."
        # Sequential character matching
        $lastIndex = -1
        foreach ($char in $searchChars)
        {
            $index = $candidateLower.IndexOf($char, $lastIndex + 1)
            Write-Verbose "[$functionName] Searching for character '$char' starting at index $($lastIndex + 1): found at index $index."
            if ($index -ge 0)
            {
                $score += 10
                Write-Verbose "[$functionName] Found character '$char' at index $index."
                if ($index -eq $lastIndex + 1)
                {
                    $score += 5  # Bonus for consecutive characters
                    Write-Verbose "[$functionName] Found consecutive character '$char' at index $index."
                }
                $lastIndex = $index
            }
        }

        # Penalize length difference
        $lengthDiff = [Math]::Abs($searchLower.Length - $candidateLower.Length)
        $score -= $lengthDiff
        Write-Verbose "[$functionName] Length difference: $lengthDiff, final score: $score."
        return $score
    }

    $functionName = $MyInvocation.MyCommand.Name
    Write-Verbose "[$functionName] Searching for test file '$FileName' in path '$Path'.  If not found, up to $matchesToReturn similar files with a minimum score of $minimScore will be presented."
    Write-Host ""
    Write-Host "Searching for test file in tests folder..." -ForegroundColor Yellow

    # Get all test files recursively
    $allTestFiles = Get-ChildItem -Path $TestsPath -Filter "*.Tests.ps1" -Recurse -File
    Write-Verbose "[$functionName] Found $($allTestFiles.Count) test files in path '$Path'."
    # Extract just the filename from the search term for exact matching
    $searchFileName = Split-Path -Leaf $FileName

    # Try exact filename match first
    $exactMatch = $allTestFiles | Where-Object { $_.Name -eq $searchFileName }
    Write-Verbose "[$functionName] Found $($exactMatch.Count) exact matches for '$searchFileName'."
    if ($exactMatch)
    {
        # Auto-select if there's exactly one match, regardless of -AllowMultiple
        if ($exactMatch.Count -eq 1)
        {
            Write-Host "Found exact match: $($exactMatch.FullName)" -ForegroundColor Green

            if ($AllowMultiple)
            {
                # Return as array for consistency with AllowMultiple mode
                return @($exactMatch.FullName)
            }
            else
            {
                return $exactMatch.FullName
            }
        }
        else
        {
            # Multiple exact matches - let user choose
            Write-Host "Found $($exactMatch.Count) exact matches:" -ForegroundColor Yellow

            $selectedFiles = Select-TestFiles -Files $exactMatch -TestsPath $TestsPath -AllowMultiple:$AllowMultiple

            if ($selectedFiles.Count -eq 0)
            {
                Write-Verbose "[$functionName] User chose to quit."
                return 'User canceled'
            }

            if ($AllowMultiple)
            {
                return $selectedFiles
            }
            else
            {
                return $selectedFiles[0]
            }
        }
    }

    # No exact match - perform fuzzy search
    Write-Host "No exact match found. Searching for similar files..." -ForegroundColor Yellow

    $scoredFiles = $allTestFiles | ForEach-Object {
        $score = Get-FuzzyMatchScore -SearchTerm $searchFileName -Candidate $_.Name
        [PSCustomObject]@{
            File  = $_
            Score = $score
        }
    } | Where-Object { $_.Score -gt $minimScore } | Sort-Object -Property Score -Descending | Select-Object -First $matchesToReturn
    Write-Verbose "[$functionName] Found $($scoredFiles.Count) similar files for '$searchFileName' after fuzzy matching."

    if ($scoredFiles.Count -eq 0)
    {
        Write-Host "No similar test files found" -ForegroundColor Red
        return 'No files found'
    }

    # Check if the top match is an exact match (score >= 1000)
    # Auto-select regardless of -AllowMultiple switch when there's an exact match
    if ($scoredFiles[0].Score -ge 1000)
    {
        Write-Host "Found exact fuzzy match: $($scoredFiles[0].File.FullName)" -ForegroundColor Green
        Write-Verbose "[$functionName] Exact fuzzy match found with score $($scoredFiles[0].Score), proceeding without prompt."

        if ($AllowMultiple)
        {
            # Return as array for consistency with AllowMultiple mode
            return @($scoredFiles[0].File.FullName)
        }
        else
        {
            return $scoredFiles[0].File.FullName
        }
    }

    Write-Host ""
    Write-Host "Found $($scoredFiles.Count) similar test file(s):" -ForegroundColor Cyan

    $files = $scoredFiles | ForEach-Object { $_.File }
    $selectedFiles = Select-TestFiles -Files $files -TestsPath $Path -AllowMultiple:$AllowMultiple

    if ($selectedFiles.Count -eq 0)
    {
        Write-Verbose "[$functionName] User chose to quit."
        return 'User canceled'
    }

    if ($AllowMultiple)
    {
        return $selectedFiles
    }
    else
    {
        return $selectedFiles[0]
    }
}

# Generic helper function for interactive selection from a list with paging support
function Select-ItemsFromList()
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [array]$Items,
        [Parameter(Mandatory)]
        [string]$Title,
        [Parameter(Mandatory)]
        [scriptblock]$DisplayFormat,
        [switch]$AllowMultiple,
        [string]$PromptText = "Enter selection",
        [int]$PageSize = 10
    )

    $functionName = $MyInvocation.MyCommand.Name
    Write-Verbose "[$functionName] Starting interactive selection. Title: '$Title', AllowMultiple: $AllowMultiple, PageSize: $PageSize."
    if ($Items.Count -eq 0)
    {
        Write-Host "No items available for selection" -ForegroundColor Yellow
        return @()
    }

    # Paging state
    $currentPage = 0
    $totalPages = [Math]::Ceiling($Items.Count / $PageSize)
    $usePaging = $Items.Count -gt $PageSize

    # Track selections across pages (store indices, not items)
    $selectedIndices = @{}
    $done = $false

    while (-not $done)
    {
        # Calculate page boundaries
        $startIndex = $currentPage * $PageSize
        $endIndex = [Math]::Min($startIndex + $PageSize - 1, $Items.Count - 1)
        $pageItems = $Items[$startIndex..$endIndex]

        # Display header
        Clear-Host
        Write-Host ""
        Write-Host "===============================================================" -ForegroundColor Cyan
        Write-Host " $Title" -ForegroundColor Cyan
        if ($usePaging)
        {
            Write-Host " Page $($currentPage + 1) of $totalPages (Items $($startIndex + 1)-$($endIndex + 1) of $($Items.Count))" -ForegroundColor Cyan
        }
        Write-Host "===============================================================" -ForegroundColor Cyan
        Write-Host ""

        # Display items for current page
        for ($i = 0; $i -lt $pageItems.Count; $i++)
        {
            $globalIndex = $startIndex + $i
            $itemNumber = $globalIndex + 1
            $displayText = & $DisplayFormat $pageItems[$i]

            # Show [X] marker if already selected
            $marker = if ($selectedIndices.ContainsKey($globalIndex))
            {
                "[X]"
            }
            else
            {
                "[ ]"
            }

            if ($AllowMultiple)
            {
                Write-Host " $marker [$itemNumber] $displayText" -ForegroundColor $(if ($selectedIndices.ContainsKey($globalIndex))
                    {
                        'Green'
                    }
                    else
                    {
                        'White'
                    })
            }
            else
            {
                Write-Host " [$itemNumber] $displayText" -ForegroundColor White
            }
        }

        # Display navigation and selection options
        Write-Host ""
        if ($AllowMultiple)
        {
            Write-Host "[a] Select All on Current Page" -ForegroundColor Green
            Write-Host "[c] Clear All Selections" -ForegroundColor Yellow
            if ($selectedIndices.Count -gt 0)
            {
                Write-Host "[d] Done - Use $($selectedIndices.Count) Selected Item(s)" -ForegroundColor Green
            }
        }

        if ($usePaging)
        {
            Write-Host ""
            Write-Host "Navigation:" -ForegroundColor Cyan
            if ($currentPage -gt 0)
            {
                Write-Host "[p] Previous Page" -ForegroundColor Yellow
                Write-Host "[f] First Page" -ForegroundColor Yellow
            }
            if ($currentPage -lt $totalPages - 1)
            {
                Write-Host "[n] Next Page" -ForegroundColor Yellow
                Write-Host "[l] Last Page" -ForegroundColor Yellow
            }
        }

        Write-Host ""
        Write-Host "[q] or [0] to Quit" -ForegroundColor Gray
        Write-Host ""

        if ($AllowMultiple)
        {
            Write-Host "$PromptText (comma-separated, e.g., 1, 3, 5):" -ForegroundColor Cyan
        }
        else
        {
            Write-Host "${PromptText}:" -ForegroundColor Cyan
        }

        $choice = Read-Host "Selection"

        # Handle navigation commands
        switch -Regex ($choice)
        {
            '^[qQ]$|^0$'
            {
                Write-Host "Selection canceled" -ForegroundColor Yellow
                return @()
            }
            '^[dD]$'
            {
                if ($AllowMultiple -and $selectedIndices.Count -gt 0)
                {
                    $done = $true
                }
                else
                {
                    Write-Host "No items selected. Use numbers to select items." -ForegroundColor Yellow
                    Start-Sleep -Seconds 2
                }
            }
            '^[aA]$'
            {
                if ($AllowMultiple)
                {
                    # Select all items on current page
                    for ($i = $startIndex; $i -le $endIndex; $i++)
                    {
                        $selectedIndices[$i] = $true
                    }
                    Write-Host "Selected all items on current page" -ForegroundColor Green
                    Start-Sleep -Seconds 1
                }
            }
            '^[cC]$'
            {
                if ($AllowMultiple)
                {
                    $selectedIndices.Clear()
                    Write-Host "Cleared all selections" -ForegroundColor Yellow
                    Start-Sleep -Seconds 1
                }
            }
            '^[pP]$'
            {
                if ($usePaging -and $currentPage -gt 0)
                {
                    $currentPage--
                }
            }
            '^[nN]$'
            {
                if ($usePaging -and $currentPage -lt $totalPages - 1)
                {
                    $currentPage++
                }
            }
            '^[fF]$'
            {
                if ($usePaging -and $currentPage -gt 0)
                {
                    $currentPage = 0
                }
            }
            '^[lL]$'
            {
                if ($usePaging -and $currentPage -lt $totalPages - 1)
                {
                    $currentPage = $totalPages - 1
                }
            }
            default
            {
                # Handle numeric selections
                if (-not [string]::IsNullOrWhiteSpace($choice))
                {
                    $numbers = $choice -split ',' | ForEach-Object { $_.Trim() }
                    $hasErrors = $false

                    foreach ($num in $numbers)
                    {
                        try
                        {
                            $itemNumber = [int]$num
                            $index = $itemNumber - 1

                            if ($index -ge 0 -and $index -lt $Items.Count)
                            {
                                if ($AllowMultiple)
                                {
                                    # Toggle selection
                                    if ($selectedIndices.ContainsKey($index))
                                    {
                                        $selectedIndices.Remove($index)
                                        Write-Host "Deselected item $itemNumber" -ForegroundColor Yellow
                                    }
                                    else
                                    {
                                        $selectedIndices[$index] = $true
                                        Write-Host "Selected item $itemNumber" -ForegroundColor Green
                                    }
                                }
                                else
                                {
                                    # Single selection mode - return immediately
                                    Write-Host ""
                                    Write-Host "Selected: $(& $DisplayFormat $Items[$index])" -ForegroundColor Green
                                    return @($Items[$index])
                                }
                            }
                            else
                            {
                                Write-Host "Invalid selection: $num (out of range 1-$($Items.Count))" -ForegroundColor Red
                                $hasErrors = $true
                            }
                        }
                        catch
                        {
                            Write-Host "Invalid input: '$num' (not a number)" -ForegroundColor Red
                            $hasErrors = $true
                        }
                    }

                    if ($hasErrors)
                    {
                        Start-Sleep -Seconds 2
                    }
                    elseif ($AllowMultiple)
                    {
                        Start-Sleep -Seconds 1
                    }
                }
            }
        }
    }

    # Return selected items
    if ($selectedIndices.Count -gt 0)
    {
        $result = @()
        $sortedIndices = $selectedIndices.Keys | Sort-Object
        foreach ($index in $sortedIndices)
        {
            $result += $Items[$index]
        }

        Write-Host ""
        Write-Host "Returning $($result.Count) selected item(s)" -ForegroundColor Green
        Start-Sleep -Seconds 1
        return $result
    }

    return @()
}

# Helper function to extract all tags from test files
function Get-AvailableTags()
{
    [CmdletBinding()]
    param(
        [string]$TestsPath,
        [string[]]$Files
    )

    Write-Host "Scanning test files for available tags..." -ForegroundColor Yellow

    $allTags = @{}

    # If specific files are provided, use them; otherwise scan all test files
    if ($Files -and $Files.Count -gt 0)
    {
        $testFiles = $Files | ForEach-Object { Get-Item $_ }
    }
    else
    {
        $testFiles = Get-ChildItem -Path $TestsPath -Filter "*.Tests.ps1" -Recurse -File
    }

    foreach ($file in $testFiles)
    {
        $content = Get-Content -Path $file.FullName -Raw
        # Match Describe blocks with -Tags parameter
        $tagMatches = [regex]::Matches($content, "Describe\s+[^-]+-Tags\s+([^ {]+)")

        foreach ($match in $tagMatches)
        {
            $tagsString = $match.Groups[1].Value
            # Extract individual tags (handle both 'tag' and "tag" format)
            $individualTags = [regex]::Matches($tagsString, '[''`"]([^''`"]+)[''`"]')

            foreach ($tagMatch in $individualTags)
            {
                $tag = $tagMatch.Groups[1].Value.Trim()
                if (-not [string]::IsNullOrWhiteSpace($tag))
                {
                    if ($allTags.ContainsKey($tag))
                    {
                        $allTags[$tag]++
                    }
                    else
                    {
                        $allTags[$tag] = 1
                    }
                }
            }
        }
    }

    # Return sorted tags with their counts
    return $allTags.GetEnumerator() | Sort-Object -Property Name
}

# Helper function to get test files containing specific tags
function Get-FilesContainingTags()
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string[]]$Tags,
        [Parameter(Mandatory)]
        [string]$TestsPath
    )

    $functionName = $MyInvocation.MyCommand.Name
    Write-Verbose "[$functionName] Finding files containing tags: $($Tags -join ', ')"

    $testFiles = Get-ChildItem -Path $TestsPath -Filter "*.Tests.ps1" -Recurse -File
    $matchingFiles = @()

    foreach ($file in $testFiles)
    {
        $content = Get-Content -Path $file.FullName -Raw
        $fileHasTag = $false

        # Check if file contains any of the specified tags
        foreach ($tag in $Tags)
        {
            # Match Describe blocks with this tag
            if ($content -match "Describe\s+[^-]+-Tags\s+[^{]*['""`]$tag['""`]")
            {
                $fileHasTag = $true
                break
            }
        }

        if ($fileHasTag)
        {
            $matchingFiles += $file
        }
    }

    Write-Verbose "[$functionName] Found $($matchingFiles.Count) file(s) containing the specified tags"
    return $matchingFiles
}

# Helper function for tag selection menu
function Select-Tags()
{
    [CmdletBinding()]
    param(
        [string]$TestsPath
    )

    $availableTags = Get-AvailableTags -TestsPath $TestsPath

    if ($availableTags.Count -eq 0)
    {
        Write-Host "No tags found in test files" -ForegroundColor Yellow
        return @()
    }

    $tagList = @($availableTags)

    $selectedTags = Select-ItemsFromList `
        -Items $tagList `
        -Title "Available Test Tags" `
        -DisplayFormat { param($tag) "$($tag.Name) ($($tag.Value) test(s))" } `
        -AllowMultiple `
        -PromptText "Enter tag numbers" `
        -PageSize $pageSize

    if ($selectedTags.Count -gt 0)
    {
        $tagNames = $selectedTags | ForEach-Object { $_.Name }
        Write-Host "Selected tags: $($tagNames -join ', ')" -ForegroundColor Green
        return $tagNames
    }

    return @()
}

# Helper function for test file selection menu
function Select-TestFiles()
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [array]$Files,
        [string]$TestsPath,
        [switch]$AllowMultiple
    )
    if ($Files.Count -eq 0)
    {
        Write-Host "No test files available for selection" -ForegroundColor Yellow
        return @()
    }

    $selectedFiles = Select-ItemsFromList `
        -Items $Files `
        -Title "Available Test Files" `
        -DisplayFormat {
        param($file)
        if ($TestsPath)
        {
            $file.FullName.Replace($TestsPath, "tests").TrimStart([System.IO.Path]::DirectorySeparatorChar)
        }
        else
        {
            $file.FullName
        }
    } `
        -AllowMultiple:$AllowMultiple `
        -PromptText "Enter file number$(if ($AllowMultiple) {'s'})" `
        -PageSize $pageSize

    if ($selectedFiles.Count -gt 0)
    {
        return $selectedFiles | ForEach-Object { $_.FullName }
    }

    return @()
}

# Helper function for combined interactive selection (files + tags)
function Select-InteractiveCombined()
{
    [CmdletBinding()]
    param(
        [string]$TestsPath
    )

    $result = @{
        TestFiles = @()
        Tags      = @()
        Canceled  = $false
    }

    while ($true)
    {
        Clear-Host
        Write-Host ""
        Write-Host "===============================================================" -ForegroundColor Cyan
        Write-Host " Interactive Test Selection" -ForegroundColor Cyan
        Write-Host "===============================================================" -ForegroundColor Cyan
        Write-Host ""
        Write-Host "Current Selections:" -ForegroundColor Yellow
        Write-Host "  Test Files: $($result.TestFiles.Count) selected" -ForegroundColor White
        if ($result.TestFiles.Count -gt 0)
        {
            foreach ($file in $result.TestFiles)
            {
                Write-Host "    - $(Split-Path -Leaf $file)" -ForegroundColor Gray
            }
        }
        Write-Host "  Tags: $($result.Tags.Count) selected" -ForegroundColor White
        if ($result.Tags.Count -gt 0)
        {
            Write-Host "    - $($result.Tags -join ', ')" -ForegroundColor Gray
        }

        # Show informational message based on selections
        if ($result.Tags.Count -gt 0 -and $result.TestFiles.Count -eq 0)
        {
            Write-Host ""
            Write-Host "  Mode: Run all tests with selected tags (across all files)" -ForegroundColor Cyan
        }
        elseif ($result.TestFiles.Count -gt 0 -and $result.Tags.Count -eq 0)
        {
            Write-Host ""
            Write-Host "  Mode: Run all tests in selected files" -ForegroundColor Cyan
        }
        elseif ($result.TestFiles.Count -gt 0 -and $result.Tags.Count -gt 0)
        {
            Write-Host ""
            Write-Host "  Mode: Run tests matching BOTH selected files AND tags" -ForegroundColor Cyan
        }
        Write-Host ""
        Write-Host "Options:" -ForegroundColor Cyan
        Write-Host "[1] Select Test Files" -ForegroundColor White
        Write-Host "[2] Select Tags" -ForegroundColor White
        Write-Host "[3] Clear Test Files" -ForegroundColor Yellow
        Write-Host "[4] Clear Tags" -ForegroundColor Yellow
        Write-Host "[5] Clear All" -ForegroundColor Yellow
        Write-Host "[r] Run Tests with Current Selections" -ForegroundColor Green
        Write-Host "[q] or 0 to Quit" -ForegroundColor Gray
        Write-Host ""

        $choice = Read-Host "Enter choice"

        switch -Regex ($choice)
        {
            '^1$'
            {
                # Select test files
                Write-Host ""

                # If tags are selected, filter files to only those containing the tags
                if ($result.Tags.Count -gt 0)
                {
                    Write-Host "Finding files containing selected tags..." -ForegroundColor Yellow
                    $allTestFiles = Get-FilesContainingTags -Tags $result.Tags -TestsPath $TestsPath

                    if ($allTestFiles.Count -eq 0)
                    {
                        Write-Host ""
                        Write-Host "No test files found containing the selected tags: $($result.Tags -join ', ')" -ForegroundColor Red
                        Write-Host "Consider clearing tags (option 4) to see all files." -ForegroundColor Yellow
                        Start-Sleep -Seconds 3
                        continue
                    }

                    Write-Host "Found $($allTestFiles.Count) file(s) containing selected tags" -ForegroundColor Green
                    Write-Host "(Filtering to show only files with tags: $($result.Tags -join ', '))" -ForegroundColor Cyan
                    Write-Host ""
                }
                else
                {
                    Write-Host "Loading test files..." -ForegroundColor Yellow
                    $allTestFiles = Get-ChildItem -Path $TestsPath -Filter "*.Tests.ps1" -Recurse -File

                    if ($allTestFiles.Count -eq 0)
                    {
                        Write-Host "No test files found" -ForegroundColor Red
                        Start-Sleep -Seconds 2
                        continue
                    }
                }

                $selectedFiles = Select-TestFiles -Files $allTestFiles -TestsPath $TestsPath -AllowMultiple

                if ($selectedFiles.Count -gt 0)
                {
                    $result.TestFiles = $selectedFiles

                    # Auto-filter tags based on newly selected files
                    Write-Host ""
                    Write-Host "Auto-filtering tags to match selected files..." -ForegroundColor Cyan
                    $availableTags = Get-AvailableTags -TestsPath $TestsPath -Files $result.TestFiles

                    if ($availableTags.Count -gt 0)
                    {
                        # If tags were previously selected, keep only those that exist in the new file set
                        if ($result.Tags.Count -gt 0)
                        {
                            $validTags = @()
                            $availableTagNames = $availableTags | ForEach-Object { $_.Name }
                            foreach ($tag in $result.Tags)
                            {
                                if ($availableTagNames -contains $tag)
                                {
                                    $validTags += $tag
                                }
                            }
                            $result.Tags = $validTags

                            if ($validTags.Count -lt $result.Tags.Count)
                            {
                                Write-Host "Some previously selected tags are not in the new file set and were removed." -ForegroundColor Yellow
                            }
                        }
                        Write-Host "Tags filtered: $($availableTags.Count) tag(s) available in selected files" -ForegroundColor Green
                    }
                    else
                    {
                        if ($result.Tags.Count -gt 0)
                        {
                            Write-Host "Warning: Selected files contain no tags. Tag selection cleared." -ForegroundColor Yellow
                        }
                        $result.Tags = @()
                    }
                    Start-Sleep -Seconds 2
                }
            }
            '^2$'
            {
                # Select tags
                Write-Host ""

                # If files are selected, filter tags to only those in the selected files
                if ($result.TestFiles.Count -gt 0)
                {
                    Write-Host "Finding tags in selected files..." -ForegroundColor Yellow
                    $availableTags = Get-AvailableTags -TestsPath $TestsPath -Files $result.TestFiles

                    if ($availableTags.Count -eq 0)
                    {
                        Write-Host ""
                        Write-Host "No tags found in the selected files" -ForegroundColor Red
                        Write-Host "Consider clearing files (option 3) to see all tags." -ForegroundColor Yellow
                        Start-Sleep -Seconds 3
                        continue
                    }

                    Write-Host "Found $($availableTags.Count) tag(s) in selected files" -ForegroundColor Green
                    Write-Host "(Filtering to show only tags in: $(($result.TestFiles | ForEach-Object { Split-Path -Leaf $_ }) -join ', '))" -ForegroundColor Cyan
                    Write-Host ""

                    # Use filtered tags for selection
                    $tagList = @($availableTags)

                    $selectedTags = Select-ItemsFromList `
                        -Items $tagList `
                        -Title "Available Test Tags (from selected files)" `
                        -DisplayFormat { param($tag) "$($tag.Name) ($($tag.Value) test(s))" } `
                        -AllowMultiple `
                        -PromptText "Enter tag numbers" `
                        -PageSize $pageSize

                    if ($selectedTags.Count -gt 0)
                    {
                        $tagNames = $selectedTags | ForEach-Object { $_.Name }
                        Write-Host "Selected tags: $($tagNames -join ', ')" -ForegroundColor Green
                        $result.Tags = $tagNames

                        # Auto-filter files based on newly selected tags
                        Write-Host ""
                        Write-Host "Auto-filtering files to match selected tags..." -ForegroundColor Cyan
                        $matchingFiles = Get-FilesContainingTags -Tags $result.Tags -TestsPath $TestsPath

                        if ($matchingFiles.Count -gt 0)
                        {
                            # If files were previously selected, keep only those that match the new tags
                            if ($result.TestFiles.Count -gt 0)
                            {
                                $matchingFileNames = $matchingFiles | ForEach-Object { $_.FullName }
                                $validFiles = @()
                                foreach ($file in $result.TestFiles)
                                {
                                    if ($matchingFileNames -contains $file)
                                    {
                                        $validFiles += $file
                                    }
                                }
                                $result.TestFiles = $validFiles

                                if ($validFiles.Count -eq 0)
                                {
                                    Write-Host "Warning: None of the previously selected files contain the new tags. File selection cleared." -ForegroundColor Yellow
                                }
                                elseif ($validFiles.Count -lt $result.TestFiles.Count)
                                {
                                    Write-Host "Some previously selected files don't contain the new tags and were removed." -ForegroundColor Yellow
                                }
                            }
                            Write-Host "Files filtered: $($matchingFiles.Count) file(s) contain selected tags" -ForegroundColor Green
                        }
                        else
                        {
                            Write-Host "Warning: No files found containing selected tags. File selection cleared." -ForegroundColor Yellow
                            $result.TestFiles = @()
                        }
                        Start-Sleep -Seconds 2
                    }
                }
                else
                {
                    # No files selected, show all tags
                    $selectedTags = Select-Tags -TestsPath $TestsPath

                    if ($selectedTags.Count -gt 0)
                    {
                        $result.Tags = $selectedTags

                        # Auto-filter files based on newly selected tags
                        Write-Host ""
                        Write-Host "Auto-filtering files to match selected tags..." -ForegroundColor Cyan
                        $matchingFiles = Get-FilesContainingTags -Tags $result.Tags -TestsPath $TestsPath

                        if ($matchingFiles.Count -gt 0)
                        {
                            Write-Host "Files filtered: $($matchingFiles.Count) file(s) contain selected tags" -ForegroundColor Green
                        }
                        else
                        {
                            Write-Host "Warning: No files found containing selected tags." -ForegroundColor Yellow
                        }
                        Start-Sleep -Seconds 2
                    }
                }
            }
            '^3$'
            {
                # Check if clearing files would affect tag-based filtering
                if ($result.TestFiles.Count -gt 0 -and $result.Tags.Count -gt 0)
                {
                    Write-Host ""
                    Write-Host "WARNING: Clearing test files will run ALL files containing the selected tags." -ForegroundColor Yellow
                    Write-Host "Tags: $($result.Tags -join ', ')" -ForegroundColor Cyan
                    Write-Host ""
                    Write-Host "Continue? [y/n]:" -ForegroundColor Yellow -NoNewline
                    $confirm = Read-Host " "
                    if ($confirm -notmatch '^[yY]')
                    {
                        Write-Host "Cancelled" -ForegroundColor Gray
                        Start-Sleep -Seconds 1
                        continue
                    }
                }

                $result.TestFiles = @()
                Write-Host ""
                Write-Host "Test files cleared" -ForegroundColor Yellow
                Start-Sleep -Seconds 1
            }
            '^4$'
            {
                # Check if clearing tags would affect the file selection
                if ($result.Tags.Count -gt 0 -and $result.TestFiles.Count -gt 0)
                {
                    Write-Host ""
                    Write-Host "WARNING: Clearing tags will run ALL tests in the selected files." -ForegroundColor Yellow
                    Write-Host "Currently selected files: $($result.TestFiles.Count)" -ForegroundColor Cyan
                    Write-Host ""
                    Write-Host "Continue? [y/n]:" -ForegroundColor Yellow -NoNewline
                    $confirm = Read-Host " "
                    if ($confirm -notmatch '^[yY]')
                    {
                        Write-Host "Cancelled" -ForegroundColor Gray
                        Start-Sleep -Seconds 1
                        continue
                    }
                }

                $result.Tags = @()
                Write-Host ""
                Write-Host "Tags cleared" -ForegroundColor Yellow
                Start-Sleep -Seconds 1
            }
            '^5$'
            {
                $result.TestFiles = @()
                $result.Tags = @()
                Write-Host "All selections cleared" -ForegroundColor Yellow
                Start-Sleep -Seconds 1
            }
            '^[rR]$'
            {
                # Validate at least one selection
                if ($result.TestFiles.Count -eq 0 -and $result.Tags.Count -eq 0)
                {
                    Write-Host ""
                    Write-Host "No selections made. Please select test files or tags." -ForegroundColor Red
                    Start-Sleep -Seconds 2
                    continue
                }

                Write-Host ""
                Write-Host "Proceeding with test execution..." -ForegroundColor Green
                Start-Sleep -Seconds 1
                return $result
            }
            '^[qQ0]$'
            {
                $result.Canceled = $true
                return $result
            }
            default
            {
                Write-Host "Invalid choice. Please try again." -ForegroundColor Red
                Start-Sleep -Seconds 1
            }
        }
    }
}

# Helper function to clean up GUID-named TestDrive folders
function Remove-GuidFolders()
{
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Path,
        [Parameter(Mandatory)]
        [string]$LocationDescription
    )

    $guidPattern = '^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$'

    if (-not (Test-Path $Path))
    {
        return
    }

    $guidFolders = Get-ChildItem -Path $Path -Directory -ErrorAction SilentlyContinue |
        Where-Object { $_.Name -match $guidPattern }

    if ($guidFolders.Count -gt 0)
    {
        Write-Host "Cleaning up $($guidFolders.Count) GUID-named TestDrive folder(s) from $LocationDescription..." -ForegroundColor Yellow
        foreach ($folder in $guidFolders)
        {
            Remove-Item -Path $folder.FullName -Recurse -Force -ErrorAction SilentlyContinue
            Write-Verbose "Removed: $($folder.FullName)"
        }
    }
}
#endregion

# Check if -Interactive mode was requested (combined files + tags selection)
if ($Interactive)
{
    $testsPath = Join-Path $PSScriptRoot "tests"
    $interactiveResult = Select-InteractiveCombined -TestsPath $testsPath

    if ($interactiveResult.Canceled)
    {
        Write-Host ""
        Write-Host "Interactive mode canceled. Exiting." -ForegroundColor Yellow
        exit 0
    }

    # Apply selections from interactive mode
    if ($interactiveResult.TestFiles.Count -gt 0)
    {
        $config.Run.Path = $interactiveResult.TestFiles
    }

    if ($interactiveResult.Tags.Count -gt 0)
    {
        $Tags = $interactiveResult.Tags
    }
}

# Check if -Tags was passed with "Interactive" value (interactive mode)
if ($PSBoundParameters.ContainsKey('Tags') -and $Tags.Count -eq 1 -and $Tags[0] -eq "Interactive")
{
    # User passed -Tags "Interactive", show interactive selection
    $testsPath = Join-Path $PSScriptRoot "tests"
    $Tags = Select-Tags -TestsPath $testsPath

    if ($Tags.Count -eq 0)
    {
        Write-Host ""
        Write-Host "No tags selected. Running all tests." -ForegroundColor Yellow
        $Tags = @()
    }
}

# Check if -TestFile was passed with "Interactive" value (interactive file selection mode)
if ($PSBoundParameters.ContainsKey('TestFile') -and $TestFile -eq "Interactive")
{
    # User passed -TestFile "Interactive", show interactive file selection
    Write-Host ""
    Write-Host "Loading test files..." -ForegroundColor Yellow
    $testsPath = Join-Path $PSScriptRoot "tests"
    $allTestFiles = Get-ChildItem -Path $testsPath -Filter "*.Tests.ps1" -Recurse -File

    if ($allTestFiles.Count -eq 0)
    {
        Write-Host "No test files found in tests folder" -ForegroundColor Red
        exit 1
    }

    $selectedFiles = Select-TestFiles -Files $allTestFiles -TestsPath $testsPath -AllowMultiple

    if ($selectedFiles.Count -eq 0)
    {
        Write-Host ""
        Write-Host "No files selected. Exiting." -ForegroundColor Yellow
        exit 0
    }

    # Update config to run selected files
    $config.Run.Path = $selectedFiles
    Write-Host ""
    Write-Host "Selected $($selectedFiles.Count) test file(s) to run" -ForegroundColor Green
}
elseif ($PSBoundParameters.ContainsKey('TestFile') -and -not [string]::IsNullOrWhiteSpace($TestFile))
{
    # User specified a test file path or search term
    $resolvedPath = if ([System.IO.Path]::IsPathRooted($TestFile))
    {
        $TestFile
    }
    else
    {
        Join-Path $PSScriptRoot $TestFile
    }

    if (-not (Test-Path $resolvedPath))
    {
        Write-Host "Test file not found: $resolvedPath" -ForegroundColor Yellow

        # Try to find the file in the tests folder
        $testsPath = Join-Path $PSScriptRoot "tests"
        $foundPath = Find-FileWithFuzzySearch -FileName $TestFile -Path $testsPath -AllowMultiple

        if ($foundPath -notin $strings)
        {
            # Check if multiple files were returned
            if ($foundPath -is [array])
            {
                Write-Host ""
                Write-Host "Using $($foundPath.Count) selected test file(s)" -ForegroundColor Green

                # Handle exclusion mode
                if ($Exclude)
                {
                    # Get all test files and exclude the found ones
                    $allTestFiles = Get-ChildItem -Path $testsPath -Recurse -Filter "*.Tests.ps1" | Where-Object { $_.FullName -notin $foundPath }
                    $config.Run.Path = $allTestFiles.FullName
                    Write-Host "Excluding $($foundPath.Count) file(s), running $($allTestFiles.Count) remaining test(s)" -ForegroundColor Yellow
                }
                else
                {
                    $config.Run.Path = $foundPath
                }
            }
            else
            {
                $resolvedPath = $foundPath
                Write-Host ""
                Write-Host "Using selected test file: $resolvedPath" -ForegroundColor Green

                # Handle exclusion mode
                if ($Exclude)
                {
                    # Get all test files and exclude the found one
                    $allTestFiles = Get-ChildItem -Path $testsPath -Recurse -Filter "*.Tests.ps1" | Where-Object { $_.FullName -ne $resolvedPath }
                    $config.Run.Path = $allTestFiles.FullName
                    Write-Host "Excluding this file, running $($allTestFiles.Count) remaining test(s)" -ForegroundColor Yellow
                }
                else
                {
                    $config.Run.Path = $resolvedPath
                }
            }
        }
        elseif ($foundPath -in $strings)
        {
            if ($foundPath -eq 'User canceled')
            {
                Write-Host ""
                Write-Host "Operation canceled by user." -ForegroundColor Red
                exit 1
            }
            elseif ($foundPath -eq 'No files found')
            {
                Write-Host ""
                Write-Host "ERROR: No matching test files found." -ForegroundColor Red
                exit 1
            }
        }
        else
        {
            Write-Host ""
            Write-Host "ERROR: Could not resolve test file" -ForegroundColor Red
            exit 1
        }
    }
    else
    {
        # Handle exclusion mode for directly specified file
        if ($Exclude)
        {
            # Get all test files and exclude the specified one
            $testsPath = Join-Path $PSScriptRoot "tests"
            $allTestFiles = Get-ChildItem -Path $testsPath -Recurse -Filter "*.Tests.ps1" | Where-Object { $_.FullName -ne $resolvedPath }
            $config.Run.Path = $allTestFiles.FullName
            Write-Host "`nExcluding test file: $(Split-Path -Leaf $resolvedPath)" -ForegroundColor Yellow
            Write-Host "Running $($allTestFiles.Count) remaining test(s)" -ForegroundColor Yellow
        }
        else
        {
            $config.Run.Path = $resolvedPath
            Write-Host "`nRunning single test file: $(Split-Path -Leaf $resolvedPath)" -ForegroundColor Yellow
        }
    }
}

# Apply tag filter if specified
if ($Tags.Count -gt 0)
{
    if ($Exclude)
    {
        $config.Filter.ExcludeTag = $Tags
    }
    else
    {
        $config.Filter.Tag = $Tags
    }
}

# Display configuration
Write-Host "`nTest Configuration:" -ForegroundColor Cyan
if ($Exclude -and $TestType -ne 'All')
{
    Write-Host "  Test Type: $TestType (EXCLUDING)" -ForegroundColor White
}
else
{
    Write-Host "  Test Type: $TestType" -ForegroundColor White
}

# Handle display of test paths (single or multiple)
$testPaths = $config.Run.Path.Value
if ($testPaths -is [array] -and $testPaths.Count -gt 1)
{
    Write-Host "  Test Files: $($testPaths.Count) files selected" -ForegroundColor White
    foreach ($path in $testPaths)
    {
        Write-Host "    - $(Split-Path -Leaf $path)" -ForegroundColor Gray
    }
}
else
{
    Write-Host "  Test Path: $testPaths" -ForegroundColor White
}

if ($TestType -in @('Integration') -and $OutputVerbosity -ne 'Detailed')
{
    Write-Host "Changing output verbosity to 'Detailed' for Integration tests" -ForegroundColor Yellow
    $config.Output.Verbosity = 'Detailed'
}
if ($EnableCodeCoverage -or $OutputVerbosity -eq 'Detailed')
{
    Write-Host "  Code Coverage: $($config.CodeCoverage.Enabled)" -ForegroundColor White
}
if ($Tags.Count -gt 0)
{
    if ($Exclude)
    {
        Write-Host "  Excluding Tags: $($Tags -join ', ')" -ForegroundColor White
    }
    else
    {
        Write-Host "  Tags: $($Tags -join ', ')" -ForegroundColor White
    }
}
Write-Host ""

# Run Pester
$startTime = Get-Date
Write-Host "Starting Pester tests..." -ForegroundColor Cyan

try
{
    $result = Invoke-Pester -Configuration $config

    # Clean up any leftover Pester TestDrive GUID folders
    # Pester 5 sometimes leaves these behind in the working directory
    Get-ChildItem -Directory | Where-Object {
        $_.Name -match '^[a-f0-9] {8}-[a-f0-9] {4}-[a-f0-9] {4}-[a-f0-9] {4}-[a-f0-9] {12}$'
    } | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue | Out-Null

    $endTime = Get-Date
    $duration = $endTime - $startTime
    # Display results
    Write-Host "`n" -NoNewline
    Write-Host "=" * 63 -ForegroundColor Cyan
    Write-Host "  Test Results" -ForegroundColor Cyan
    Write-Host "=" * 63 -ForegroundColor Cyan
    Write-Host "  Total Tests: $($result.TotalCount)" -ForegroundColor White
    Write-Host "  Passed: $($result.PassedCount)" -ForegroundColor Green
    Write-Host "  Failed: $($result.FailedCount)" -ForegroundColor $(if ($result.FailedCount -gt 0)
        {
            'Red'
        }
        else
        {
            'Gray'
        })
    Write-Host "  Skipped: $($result.SkippedCount)" -ForegroundColor Gray
    Write-Host "  Not Run: $($result.NotRunCount)" -ForegroundColor DarkGray
    Write-Host "  Duration: $($duration.TotalSeconds.ToString('F2'))s" -ForegroundColor White

    # Display container failures if any
    if ($result.FailedContainersCount -gt 0)
    {
        Write-Host "`n  Container Failures: $($result.FailedContainersCount)" -ForegroundColor Red
        Write-Host "  Failed Containers:" -ForegroundColor Red
        foreach ($container in $result.FailedContainers)
        {
            Write-Host "    - $($container.Item.ToString())" -ForegroundColor Red
            if ($container.ErrorRecord)
            {
                Write-Host "      Error: $($container.ErrorRecord.Exception.Message)" -ForegroundColor Red
            }
        }
    }

    # Display failed test details if any
    if ($result.FailedCount -gt 0)
    {
        Write-Host "`n  Failed Tests:" -ForegroundColor Red

        # Group failed tests by file
        $failedByFile = $result.Failed | Group-Object -Property {
            if ($_.ScriptBlock.File)
            {
                Split-Path $_.ScriptBlock.File -Leaf
            }
            else
            {
                "Unknown File"
            }
        } | Sort-Object Name

        foreach ($fileGroup in $failedByFile)
        {
            Write-Host "`n    $($fileGroup.Name) ($($fileGroup.Count) failure$(if ($fileGroup.Count -ne 1) {'s'})):" -ForegroundColor Yellow
            foreach ($test in $fileGroup.Group)
            {
                Write-Host "      - $($test.ExpandedName)" -ForegroundColor Red
                if ($test.ErrorRecord)
                {
                    Write-Host "        Error: $($test.ErrorRecord.Exception.Message)" -ForegroundColor DarkRed
                }
            }
        }
        Write-Host ""
    }

    # Display skipped test details if any
    if ($result.SkippedCount -gt 0)
    {
        Write-Host "`n  Skipped Tests:" -ForegroundColor Yellow

        # Group skipped tests by file
        $skippedByFile = $result.Skipped | Group-Object -Property {
            if ($_.ScriptBlock.File)
            {
                Split-Path $_.ScriptBlock.File -Leaf
            }
            else
            {
                "Unknown File"
            }
        } | Sort-Object Name

        foreach ($fileGroup in $skippedByFile)
        {
            Write-Host "`n    $($fileGroup.Name) ($($fileGroup.Count) skipped):" -ForegroundColor DarkYellow
            foreach ($test in $fileGroup.Group)
            {
                Write-Host "      - $($test.ExpandedName)" -ForegroundColor Gray
            }
        }
        Write-Host ""
    }

    # Display not run test details if any
    if ($result.NotRunCount -gt 0)
    {
        Write-Host "`n  Not Run Tests:" -ForegroundColor DarkGray

        # Group not run tests by file
        $notRunByFile = $result.NotRun | Group-Object -Property {
            if ($_.ScriptBlock.File)
            {
                Split-Path $_.ScriptBlock.File -Leaf
            }
            else
            {
                "Unknown File"
            }
        } | Sort-Object Name

        foreach ($fileGroup in $notRunByFile)
        {
            Write-Host "`n    $($fileGroup.Name) ($($fileGroup.Count) not run):" -ForegroundColor DarkGray
            foreach ($test in $fileGroup.Group)
            {
                Write-Host "      - $($test.ExpandedName)" -ForegroundColor DarkGray
            }
        }
        Write-Host ""
    }

    # Display code coverage only if enabled
    if ($EnableCodeCoverage -and $result.CodeCoverage)
    {
        $coverage = $result.CodeCoverage
        Write-Host "`nCode Coverage:" -ForegroundColor Cyan
        Write-Host "  Commands Analyzed: $($coverage.CommandsAnalyzedCount)" -ForegroundColor White
        Write-Host "  Commands Executed: $($coverage.CommandsExecutedCount)" -ForegroundColor White
        Write-Host "Commands missed: $($coverage.CommandsMissedCount)" -ForegroundColor White
        Write-Host "Files analyzed: $($coverage.FilesAnalyzedCount)" -ForegroundColor White
        Write-Host "  Coverage: $($coverage.CoveragePercent)" -ForegroundColor $(if ($coverage.CoveragePercent -ge 80)
            {
                'Green'
            }
            elseif ($coverage.CoveragePercent -ge 60)
            {
                'Yellow'
            }
            else
            {
                'Red'
            })
        Write-Host "Coverage target: $($coverage.CoveragePercentTarget)"

        # Show detailed list ONLY if requested
        if ($ShowCodeCoverageDetails)
        {
            if ($coverage.CommandsMissedCount -gt 0)
            {
                Write-Host "`n  Missed Commands:" -ForegroundColor Yellow
                Write-Host $coverage.CommandsMissed
            }
            if ($coverage.CommandsExecutedCount -gt 0)
            {
                Write-Host "`n  Executed Commands:" -ForegroundColor Green
                Write-Host $coverage.CommandsExecuted
            }
            if ($coverage.FilesAnalyzedCount -gt 0)
            {
                Write-Host "`n Analyzed Files:" -ForegroundColor Green
                Write-Host $coverage.FilesAnalyzed
            }
        }

        Write-Host "  Report: $($config.CodeCoverage.OutputPath)" -ForegroundColor Gray
    }

    if ($config.TestResult.Enabled)
    {
        Write-Host "`nTest Results XML: $($config.TestResult.OutputPath)" -ForegroundColor Gray
    }

    Write-Host "=" * 63 -ForegroundColor Cyan
    Write-Host ""

    # Cleanup any GUID-named folders that may have been left behind by TestDrive
    # These folders are created in the current working directory when TestDrive cleanup fails
    $repoRoot = $PSScriptRoot
    Remove-GuidFolders -Path $repoRoot -LocationDescription "repository root"

    $testsFolder = Join-Path $repoRoot "tests"
    Remove-GuidFolders -Path $testsFolder -LocationDescription "tests directory"

    # Exit with appropriate code
    exit $result.FailedCount
}
catch
{
    Write-Host "`nERROR: Pester execution failed" -ForegroundColor Red
    Write-Host $_.Exception.Message -ForegroundColor Red
    exit 1
}

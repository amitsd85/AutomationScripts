# Combined SharePoint PreCheck and Report Processing Script
# This script performs SharePoint pre-checks and processes the exported reports

param(
    [Parameter(Mandatory=$false)]
    [string]$ConfigFile = ".\config.json"
)

# Initialize logging
$LogFile = "C:\Logs\SharePointPreCheck_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"
$LogDirectory = Split-Path $LogFile -Parent

# Create log directory if it doesn't exist
if (!(Test-Path $LogDirectory)) {
    New-Item -ItemType Directory -Path $LogDirectory -Force | Out-Null
}

# Logging function
function Write-Log {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        [Parameter(Mandatory=$false)]
        [ValidateSet("INFO", "WARNING", "ERROR", "SUCCESS")]
        [string]$Level = "INFO"
    )
    
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    
    # Write to log file
    Add-Content -Path $LogFile -Value $logEntry
    
    # Write to console with colors
    switch ($Level) {
        "INFO"    { Write-Host $logEntry -ForegroundColor White }
        "WARNING" { Write-Host $logEntry -ForegroundColor Yellow }
        "ERROR"   { Write-Host $logEntry -ForegroundColor Red }
        "SUCCESS" { Write-Host $logEntry -ForegroundColor Green }
    }
}

# Function to check and install required modules
function Test-And-Install-Modules {
    Write-Log "Checking required PowerShell modules..." -Level "INFO"
    
    $requiredModules = @("ShareGate", "ImportExcel")
    $modulesInstalled = $true
    
    foreach ($module in $requiredModules) {
        try {
            Write-Log "Checking module: $module" -Level "INFO"
            
            if (!(Get-Module -ListAvailable -Name $module)) {
                Write-Log "Module $module is not installed. Attempting to install..." -Level "WARNING"
                
                if ($module -eq "ShareGate") {
                    Write-Log "ShareGate requires manual installation from ShareGate website" -Level "ERROR"
                    Write-Log "Please download and install ShareGate PowerShell from: https://sharegate.com/products/sharegate-desktop" -Level "ERROR"
                    $modulesInstalled = $false
                } else {
                    Install-Module -Name $module -Force -Scope CurrentUser -AllowClobber
                    Write-Log "Successfully installed module: $module" -Level "SUCCESS"
                }
            } else {
                Write-Log "Module $module is already installed" -Level "SUCCESS"
            }
            
            # Import the module
            Import-Module $module -Force
            Write-Log "Successfully imported module: $module" -Level "SUCCESS"
            
        } catch {
            Write-Log "Failed to install/import module $module : $($_.Exception.Message)" -Level "ERROR"
            $modulesInstalled = $false
        }
    }
    
    return $modulesInstalled
}

# Function to load configuration
function Get-Configuration {
    param([string]$ConfigPath)
    
    try {
        if (Test-Path $ConfigPath) {
            $config = Get-Content $ConfigPath | ConvertFrom-Json
            Write-Log "Configuration loaded from: $ConfigPath" -Level "SUCCESS"
            return $config
        } else {
            Write-Log "Configuration file not found. Using default parameters." -Level "WARNING"
            
            # Default configuration
            $defaultConfig = @{
                csvPath = "C:\YourPath\SourceSites.csv"
                destinationUrl = "https://yourcompany.sharepoint.com/sites/destination/"
                reportPath = "C:\Reports"
                username = "DOMAIN\username"
                password = "YourPassword"  # Use SecureString in production
                consolidatedOutputFile = "C:\Output\ConsolidatedReport.xlsx"
            }
            
            return $defaultConfig
        }
    } catch {
        Write-Log "Error loading configuration: $($_.Exception.Message)" -Level "ERROR"
        throw
    }
}

# Function to perform SharePoint pre-checks
function Invoke-SharePointPreCheck {
    param(
        [object]$Config
    )
    
    Write-Log "Starting SharePoint Pre-Check process..." -Level "INFO"
    
    try {
        # Convert password to SecureString
        $securePassword = ConvertTo-SecureString $Config.password -AsPlainText -Force
        Write-Log "Password converted to SecureString" -Level "SUCCESS"
        
        # Create report directory if it doesn't exist
        if (!(Test-Path $Config.reportPath)) {
            New-Item -ItemType Directory -Path $Config.reportPath -Force | Out-Null
            Write-Log "Created report directory: $($Config.reportPath)" -Level "SUCCESS"
        }
        
        # Import CSV with site URLs
        if (!(Test-Path $Config.csvPath)) {
            Write-Log "CSV file not found: $($Config.csvPath)" -Level "ERROR"
            throw "CSV file not found: $($Config.csvPath)"
        }
        
        $sites = Import-Csv -Path $Config.csvPath
        Write-Log "Loaded $($sites.Count) sites from CSV file" -Level "SUCCESS"
        
        # Connect to destination site
        Write-Log "Connecting to destination site: $($Config.destinationUrl)" -Level "INFO"
        $dstSite = Connect-Site -Url $Config.destinationUrl -Browser
        Write-Log "Successfully connected to destination site" -Level "SUCCESS"
        
        $processedSites = 0
        $successfulChecks = 0
        $failedChecks = 0
        $reportsGenerated = @()
        
        foreach ($site in $sites) {
            $srcUrl = $site.SiteUrl
            $processedSites++
            
            Write-Log "Processing site $processedSites of $($sites.Count): $srcUrl" -Level "INFO"
            
            try {
                # Connect to source site
                $srcSite = Connect-Site -Url $srcUrl -Username $Config.username -Password $securePassword
                Write-Log "Connected to source site: $srcUrl" -Level "SUCCESS"
                
                # Perform pre-check
                Write-Log "Running pre-check for: $srcUrl" -Level "INFO"
                $report = Copy-Site -Site $srcSite -DestinationSite $dstSite -WhatIf
                
                # Check if there are warnings or errors
                if (($report.Warnings -gt 0) -or ($report.Errors -gt 0)) {
                    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
                    $siteTitleSafe = ($srcSite.Title -replace "[^a-zA-Z0-9]", "_")
                    $fileName = "${siteTitleSafe}_PreCheck_$timestamp.xlsx"
                    $fullPath = Join-Path $Config.reportPath $fileName
                    
                    # Export report
                    Export-Report $report -Path $fullPath
                    $reportsGenerated += $fullPath
                    
                    Write-Log "Report exported: $fullPath (Warnings: $($report.Warnings), Errors: $($report.Errors))" -Level "SUCCESS"
                } else {
                    Write-Log "No warnings or errors found for: $srcUrl" -Level "SUCCESS"
                }
                
                $successfulChecks++
                
            } catch {
                Write-Log "Failed to process site $srcUrl : $($_.Exception.Message)" -Level "ERROR"
                $failedChecks++
            }
        }
        
        Write-Log "Pre-check completed. Processed: $processedSites, Successful: $successfulChecks, Failed: $failedChecks" -Level "INFO"
        Write-Log "Generated $($reportsGenerated.Count) reports with warnings/errors" -Level "INFO"
        
        return $reportsGenerated
        
    } catch {
        Write-Log "Error in SharePoint Pre-Check: $($_.Exception.Message)" -Level "ERROR"
        throw
    }
}

# Function to process and consolidate reports
function Invoke-ReportProcessing {
    param(
        [string]$FolderPath,
        [string]$OutputFile
    )
    
    Write-Log "Starting report processing and consolidation..." -Level "INFO"
    
    try {
        # Create output directory if it doesn't exist
        $outputDir = Split-Path $OutputFile -Parent
        if (!(Test-Path $outputDir)) {
            New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
            Write-Log "Created output directory: $outputDir" -Level "SUCCESS"
        }
        
        # Get all Excel files in the folder
        if (!(Test-Path $FolderPath)) {
            Write-Log "Report folder not found: $FolderPath" -Level "ERROR"
            throw "Report folder not found: $FolderPath"
        }
        
        $excelFiles = Get-ChildItem -Path $FolderPath -Filter "*.xlsx"
        Write-Log "Found $($excelFiles.Count) Excel files to process" -Level "INFO"
        
        if ($excelFiles.Count -eq 0) {
            Write-Log "No Excel files found in: $FolderPath" -Level "WARNING"
            return
        }
        
        # Initialize arrays to store data
        $allData = @()
        $filteredUserWarnings = @()
        $filteredGroup = @()
        
        # Process each Excel file
        foreach ($file in $excelFiles) {
            Write-Log "Processing file: $($file.Name)" -Level "INFO"
            
            try {
                $excelData = Import-Excel -Path $file.FullName
                Write-Log "Imported $($excelData.Count) rows from $($file.Name)" -Level "SUCCESS"
                
                # Add all data to the consolidated array
                $allData += $excelData
                
                # Filter rows for User with Result = Warning or Error
                $userWarnings = $excelData | Where-Object {
                    $_.Type -eq "User" -and
                    ($_.Result -eq "Warning" -or $_.Result -eq "Error")
                }
                $filteredUserWarnings += $userWarnings
                Write-Log "Found $($userWarnings.Count) user warnings/errors in $($file.Name)" -Level "INFO"
                
                # Filter rows for Group with Result = Warning or Error
                $groupWarnings = $excelData | Where-Object {
                    $_.Type -eq "Group" -and
                    ($_.Result -eq "Warning" -or $_.Result -eq "Error")
                }
                $filteredGroup += $groupWarnings
                Write-Log "Found $($groupWarnings.Count) group warnings/errors in $($file.Name)" -Level "INFO"
                
            } catch {
                Write-Log "Error processing file $($file.Name): $($_.Exception.Message)" -Level "ERROR"
            }
        }
        
        # Export consolidated data to Excel
        Write-Log "Exporting consolidated data to: $OutputFile" -Level "INFO"
        
        try {
            # Export all data
            $allData | Export-Excel -Path $OutputFile -WorksheetName "ConsolidatedData" -AutoSize
            Write-Log "Exported $($allData.Count) total records to ConsolidatedData worksheet" -Level "SUCCESS"
            
            # Export filtered user warnings
            $filteredUserWarnings | Export-Excel -Path $OutputFile -WorksheetName "FilteredUserWarnings" -AutoSize -Append
            Write-Log "Exported $($filteredUserWarnings.Count) user warnings/errors to FilteredUserWarnings worksheet" -Level "SUCCESS"
            
            # Export filtered group warnings
            $filteredGroup | Export-Excel -Path $OutputFile -WorksheetName "FilteredGroup" -AutoSize -Append
            Write-Log "Exported $($filteredGroup.Count) group warnings/errors to FilteredGroup worksheet" -Level "SUCCESS"
            
            Write-Log "Report processing completed successfully!" -Level "SUCCESS"
            
        } catch {
            Write-Log "Error exporting consolidated report: $($_.Exception.Message)" -Level "ERROR"
            throw
        }
        
    } catch {
        Write-Log "Error in report processing: $($_.Exception.Message)" -Level "ERROR"
        throw
    }
}

# Main execution
try {
    Write-Log "========== SharePoint PreCheck and Report Processing Script Started ==========" -Level "INFO"
    
    # Step 1: Check and install required modules
    Write-Log "Step 1: Checking required modules..." -Level "INFO"
    if (!(Test-And-Install-Modules)) {
        Write-Log "Required modules are missing. Please install them manually and run the script again." -Level "ERROR"
        exit 1
    }
    
    # Step 2: Load configuration
    Write-Log "Step 2: Loading configuration..." -Level "INFO"
    $config = Get-Configuration -ConfigPath $ConfigFile
    
    # Step 3: Perform SharePoint pre-checks
    Write-Log "Step 3: Performing SharePoint pre-checks..." -Level "INFO"
    $generatedReports = Invoke-SharePointPreCheck -Config $config
    
    # Step 4: Process and consolidate reports
    Write-Log "Step 4: Processing and consolidating reports..." -Level "INFO"
    Invoke-ReportProcessing -FolderPath $config.reportPath -OutputFile $config.consolidatedOutputFile
    
    Write-Log "========== Script completed successfully! ==========" -Level "SUCCESS"
    Write-Log "Log file: $LogFile" -Level "INFO"
    Write-Log "Consolidated report: $($config.consolidatedOutputFile)" -Level "INFO"
    
    # Display configuration summary
    Write-Log "========== Configuration Summary ==========" -Level "INFO"
    Write-Log "CSV Source: $($config.csvPath)" -Level "INFO"
    Write-Log "Destination URL: $($config.destinationUrl)" -Level "INFO"
    Write-Log "Reports Path: $($config.reportPath)" -Level "INFO"
    Write-Log "Username: $($config.username)" -Level "INFO"
    Write-Log "Final Report: $($config.consolidatedOutputFile)" -Level "INFO"
    
} catch {
    Write-Log "========== Script failed with error ==========" -Level "ERROR"
    Write-Log "Error: $($_.Exception.Message)" -Level "ERROR"
    Write-Log "Stack Trace: $($_.ScriptStackTrace)" -Level "ERROR"
    exit 1
}
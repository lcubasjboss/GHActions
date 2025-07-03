# This PowerShell script creates an Excel file with two worksheets:
# 1. 'Pipeline Info' with Pipeline Name, Number of Execution, and Environment.
# 2. 'Repo Info' with Repository Name and Git SHA Short Version.
# It also handles the installation of its required 'ImportExcel' module (leveraging caching for speed).

param(
    [Parameter(Mandatory=$true)] # The name of the pipeline/workflow.
    [string]$PipelineName,

    [Parameter(Mandatory=$true)] # The GitHub Actions run number.
    [int]$RunNumber,

    [Parameter(Mandatory=$true)] # The name of the GitHub repository.
    [string]$RepoName,

    [Parameter(Mandatory=$true)] # The full Git commit SHA.
    [string]$CommitSha,

    [Parameter(Mandatory=$true)] # The chosen environment.
    [string]$Environment # <--- Added Environment parameter
)

# --- Configure ANSI Output ---
# Ensure PowerShell is configured to interpret ANSI escape sequences.
# This is typically needed for colored output in non-interactive sessions like CI/CD.
$PSStyle.OutputRendering = 'Ansi'
Write-Host "ANSI output rendering enabled."

# --- Dependency Installation ---
Write-Host "Checking for and installing required PowerShell modules..."

# Ensure PSGallery is registered as a trusted repository.
Write-Host "Checking/Registering PSGallery repository..."
try {
    if (-not (Get-PSRepository -Name PSGallery -ErrorAction SilentlyContinue)) {
        Register-PSRepository -Default -InstallationPolicy Trusted -ErrorAction Stop
        Write-Host "PSGallery repository registered and trusted."
    } else {
        Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction Stop
        Write-Host "PSGallery repository found and set to trusted."
    }
}
catch {
    Write-Error "Failed to register/trust PSGallery repository: $($_.Exception.Message)"
    exit 1 # Exit with an error code if repository setup fails
}

# List repositories to confirm (for debugging purposes in logs)
Write-Host "Available PowerShell Repositories:"
Get-PSRepository | Format-Table -AutoSize

# Function to safely install/update a PowerShell module
function Install-ModuleIfMissing {
    param(
        [Parameter(Mandatory=$true)]
        [string]$ModuleName
    )

    Write-Host "Checking module: $ModuleName"
    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        Write-Host "$ModuleName module not found. Installing..."
        try {
            # Try to remove it first in case of partial/corrupted install
            Remove-Module -Name $ModuleName -ErrorAction SilentlyContinue -Force
            Uninstall-Module -Name $ModuleName -ErrorAction SilentlyContinue -Force

            Install-Module -Name $ModuleName -Force -Scope CurrentUser -AllowClobber -Repository PSGallery -ErrorAction Stop
            Write-Host "$ModuleName installed successfully."
        }
        catch {
            Write-Error "Failed to install/update $ModuleName : $($_.Exception.Message)"
            exit 1 # Exit with an error code if installation fails
        }
    } else {
        Write-Host "$ModuleName module found."
    }
}

# Install/update required modules
Install-ModuleIfMissing -ModuleName PackageManagement
Install-ModuleIfMissing -ModuleName PowerShellGet
Install-ModuleIfMissing -ModuleName ImportExcel

# Import the module for use in the current session.
Write-Output "`x1b[34mImportExcel module loaded.`x1b[0m" # <--- Changed to Write-Output
Import-Module -Name ImportExcel -ErrorAction Stop

# Define the name and path for the output Excel file.
$excelFilePath = "pipeline_repo_info.xlsx"

Write-Output "`x1b[34mStarting Excel file creation at $excelFilePath...`x1b[0m" # <--- Changed to Write-Output

# --- Worksheet 1: Pipeline Info ---
# Prepare the data for the first worksheet, explicitly casting to PSCustomObject.
$pipelineInfoData = @(
    [PSCustomObject]@{
        "Pipeline Name"       = $PipelineName;
        "Number of Execution" = $RunNumber;
        "Environment"         = $Environment
    }
)

Write-Output "`x1b[34mCreating 'Pipeline Info' worksheet...`x1b[0m" # <--- Changed to Write-Output
$pipelineInfoData | Export-Excel -Path $excelFilePath `
                                 -WorksheetName "Pipeline Info" `
                                 -TableName "PipelineDetails" `
                                 -TableStyle Light9 `
                                 -ClearSheet `
                                 -AutoSize

# --- Worksheet 2: Repo Info ---
# Get the short version of the Git commit SHA (first 7 characters).
$shortCommitSha = $CommitSha.Substring(0, [System.Math]::Min(7, $CommitSha.Length))

# Prepare the data for the second worksheet, explicitly casting to PSCustomObject.
$repoInfoData = @(
    [PSCustomObject]@{
        "Repository Name"       = $RepoName;
        "Git SHA Short Version" = $shortCommitSha
    }
)

Write-Output "`x1b[34mCreating 'Repo Info' worksheet...`x1b[0m" # <--- Changed to Write-Output
$repoInfoData | Export-Excel -Path $excelFilePath `
                             -WorksheetName "Repo Info" `
                             -TableName "RepoDetails" `
                             -TableStyle Light9 `
                             -AutoSize `
                             -Append

Write-Output "`x1b[34mExcel file '$excelFilePath' created successfully and ready for upload.`x1b[0m" # <--- Changed to Write-Output

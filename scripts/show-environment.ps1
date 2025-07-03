# This PowerShell script accepts a single parameter: Environment.
# It's designed to demonstrate how to receive and display an input from a GitHub Actions workflow.

param(
    [Parameter(Mandatory=$true)] # Make the Environment parameter mandatory
    [string]$Environment         # Define the parameter as a string type
)

# Use Write-Host to display the value of the Environment parameter.
# This will be visible in the GitHub Actions workflow logs.
Write-Host "The selected environment is: $Environment"

# You can add more complex logic here based on the environment,
# for example, loading environment-specific configurations or performing deployments.
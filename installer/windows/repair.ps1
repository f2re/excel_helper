[CmdletBinding()]
param(
  [string]$PayloadRoot = '',
  [switch]$Silent
)

$arguments = @{
  Mode = 'Full'
  Silent = $Silent
}
if (-not [string]::IsNullOrWhiteSpace($PayloadRoot)) { $arguments.PayloadRoot = $PayloadRoot }
& (Join-Path $PSScriptRoot 'install.ps1') @arguments

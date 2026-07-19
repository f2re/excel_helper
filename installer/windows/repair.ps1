[CmdletBinding()]
param(
  [string]$PayloadRoot = (Join-Path $PSScriptRoot '..\..\release\payload'),
  [switch]$Silent
)
& (Join-Path $PSScriptRoot 'install.ps1') -PayloadRoot $PayloadRoot -Mode Full -Force -Silent:$Silent

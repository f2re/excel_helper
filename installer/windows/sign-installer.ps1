[CmdletBinding()]
param(
  [Parameter(Mandatory = $true)][string]$InstallerPath,
  [string]$CertificateThumbprint,
  [string]$PfxPath,
  [System.Security.SecureString]$PfxPassword,
  [string]$TimestampUrl = 'http://timestamp.digicert.com'
)

Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
$installer = [IO.Path]::GetFullPath($InstallerPath)
if (-not (Test-Path -LiteralPath $installer)) { throw "Installer not found: $installer" }
$signtool = Get-Command signtool.exe -ErrorAction Stop | Select-Object -ExpandProperty Source
$arguments = @('sign', '/fd', 'SHA256', '/tr', $TimestampUrl, '/td', 'SHA256')
if ($PfxPath) {
  $arguments += @('/f', [IO.Path]::GetFullPath($PfxPath))
  if ($PfxPassword) {
    $bstr = [Runtime.InteropServices.Marshal]::SecureStringToBSTR($PfxPassword)
    try { $arguments += @('/p', [Runtime.InteropServices.Marshal]::PtrToStringBSTR($bstr)) }
    finally { [Runtime.InteropServices.Marshal]::ZeroFreeBSTR($bstr) }
  }
} elseif ($CertificateThumbprint) {
  $arguments += @('/sha1', $CertificateThumbprint)
} else {
  throw 'Укажите -PfxPath или -CertificateThumbprint.'
}
$arguments += $installer
& $signtool @arguments
if ($LASTEXITCODE -ne 0) { throw "signtool failed with exit code $LASTEXITCODE" }
& $signtool verify /pa /all /v $installer
if ($LASTEXITCODE -ne 0) { throw "Signature verification failed with exit code $LASTEXITCODE" }
Write-Host "Signed installer: $installer"

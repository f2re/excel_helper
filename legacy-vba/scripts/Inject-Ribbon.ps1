Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'

function Read-ProfiZipText {
  param(
    [Parameter(Mandatory = $true)]$Archive,
    [Parameter(Mandatory = $true)][string]$EntryName
  )
  $entry = $Archive.GetEntry($EntryName)
  if ($null -eq $entry) { throw "OOXML entry not found: $EntryName" }
  $stream = $entry.Open()
  $reader = New-Object System.IO.StreamReader($stream, [Text.Encoding]::UTF8, $true)
  try { return $reader.ReadToEnd() } finally { $reader.Dispose(); $stream.Dispose() }
}

function Write-ProfiZipText {
  param(
    [Parameter(Mandatory = $true)]$Archive,
    [Parameter(Mandatory = $true)][string]$EntryName,
    [Parameter(Mandatory = $true)][string]$Content
  )
  $existing = $Archive.GetEntry($EntryName)
  if ($null -ne $existing) { $existing.Delete() }
  $entry = $Archive.CreateEntry($EntryName)
  $stream = $entry.Open()
  $writer = New-Object System.IO.StreamWriter($stream, (New-Object Text.UTF8Encoding($false)))
  try { $writer.Write($Content) } finally { $writer.Dispose(); $stream.Dispose() }
}

function Add-ProfiRibbon {
  [CmdletBinding()]
  param(
    [Parameter(Mandatory = $true)][string]$WorkbookPath,
    [Parameter(Mandatory = $true)][string]$RibbonXmlPath
  )
  $workbook = [IO.Path]::GetFullPath($WorkbookPath)
  $ribbonFile = [IO.Path]::GetFullPath($RibbonXmlPath)
  if (-not (Test-Path -LiteralPath $workbook)) { throw "Workbook not found: $workbook" }
  if (-not (Test-Path -LiteralPath $ribbonFile)) { throw "Ribbon XML not found: $ribbonFile" }

  Add-Type -AssemblyName System.IO.Compression
  Add-Type -AssemblyName System.IO.Compression.FileSystem
  $archive = [IO.Compression.ZipFile]::Open($workbook, [IO.Compression.ZipArchiveMode]::Update)
  try {
    [xml]$contentTypes = Read-ProfiZipText -Archive $archive -EntryName '[Content_Types].xml'
    $ctNs = New-Object Xml.XmlNamespaceManager($contentTypes.NameTable)
    $ctNs.AddNamespace('ct', 'http://schemas.openxmlformats.org/package/2006/content-types')
    $override = $contentTypes.SelectSingleNode("/ct:Types/ct:Override[@PartName='/customUI/customUI.xml']", $ctNs)
    if ($null -eq $override) {
      $override = $contentTypes.CreateElement('Override', 'http://schemas.openxmlformats.org/package/2006/content-types')
      $override.SetAttribute('PartName', '/customUI/customUI.xml')
      $override.SetAttribute('ContentType', 'application/vnd.ms-office.customUI+xml')
      [void]$contentTypes.DocumentElement.AppendChild($override)
    }

    [xml]$relationships = Read-ProfiZipText -Archive $archive -EntryName '_rels/.rels'
    $relNs = New-Object Xml.XmlNamespaceManager($relationships.NameTable)
    $relNs.AddNamespace('r', 'http://schemas.openxmlformats.org/package/2006/relationships')
    $relationship = $relationships.SelectSingleNode("/r:Relationships/r:Relationship[@Type='http://schemas.microsoft.com/office/2006/relationships/ui/extensibility']", $relNs)
    if ($null -eq $relationship) {
      $relationship = $relationships.CreateElement('Relationship', 'http://schemas.openxmlformats.org/package/2006/relationships')
      $relationship.SetAttribute('Id', 'rIdProfiCustomUI')
      $relationship.SetAttribute('Type', 'http://schemas.microsoft.com/office/2006/relationships/ui/extensibility')
      $relationship.SetAttribute('Target', 'customUI/customUI.xml')
      [void]$relationships.DocumentElement.AppendChild($relationship)
    } else {
      $relationship.SetAttribute('Target', 'customUI/customUI.xml')
    }

    $ribbonXml = Get-Content -LiteralPath $ribbonFile -Raw -Encoding UTF8
    Write-ProfiZipText -Archive $archive -EntryName '[Content_Types].xml' -Content $contentTypes.OuterXml
    Write-ProfiZipText -Archive $archive -EntryName '_rels/.rels' -Content $relationships.OuterXml
    Write-ProfiZipText -Archive $archive -EntryName 'customUI/customUI.xml' -Content $ribbonXml
  } finally {
    $archive.Dispose()
  }
  Write-Host "RibbonX injected: $workbook"
}

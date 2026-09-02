param(
  [string]$InputPath = "RAV-Task1.2-50states.xlsx",
  [string]$OutputPath = "data/regulations.json"
)

Add-Type -AssemblyName System.IO.Compression

$stateNames = @(
  "Alabama", "Alaska", "Arizona", "Arkansas", "California", "Colorado",
  "Connecticut", "Delaware", "Florida", "Georgia", "Hawaii", "Idaho",
  "Illinois", "Indiana", "Iowa", "Kansas", "Kentucky", "Louisiana", "Maine",
  "Maryland", "Massachusetts", "Michigan", "Minnesota", "Mississippi",
  "Missouri", "Montana", "Nebraska", "Nevada", "New Hampshire", "New Jersey",
  "New Mexico", "New York", "North Carolina", "North Dakota", "Ohio", "Oklahoma",
  "Oregon", "Pennsylvania", "Rhode Island", "South Carolina", "South Dakota",
  "Tennessee", "Texas", "Utah", "Vermont", "Virginia", "Washington",
  "West Virginia", "Wisconsin", "Wyoming"
)
$stateSet = @{}
$stateNames | ForEach-Object { $stateSet[$_] = $true }

function Get-EntryText($archive, [string]$name) {
  $entry = $archive.GetEntry($name)
  if (-not $entry) { throw "Missing workbook entry: $name" }
  $reader = [IO.StreamReader]::new($entry.Open())
  try { $reader.ReadToEnd() } finally { $reader.Dispose() }
}

function Clean-Text($value) {
  if ($null -eq $value) { return "" }
  $text = ([string]$value).Replace("`r", "").Trim()
  return [regex]::Replace($text, "`n{3,}", "`n`n")
}

$resolvedInput = (Resolve-Path -LiteralPath $InputPath).Path
$stream = [IO.File]::Open(
  $resolvedInput,
  [IO.FileMode]::Open,
  [IO.FileAccess]::Read,
  [IO.FileShare]::ReadWrite
)
$archive = [IO.Compression.ZipArchive]::new(
  $stream,
  [IO.Compression.ZipArchiveMode]::Read
)

try {
  [xml]$sharedXml = Get-EntryText $archive "xl/sharedStrings.xml"
  $sharedNs = [Xml.XmlNamespaceManager]::new($sharedXml.NameTable)
  $sharedNs.AddNamespace("s", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
  $sharedStrings = @(
    foreach ($item in $sharedXml.SelectNodes("//s:si", $sharedNs)) {
      ($item.SelectNodes(".//s:t", $sharedNs) | ForEach-Object { $_.InnerText }) -join ""
    }
  )

  [xml]$sheetXml = Get-EntryText $archive "xl/worksheets/sheet1.xml"
  $sheetNs = [Xml.XmlNamespaceManager]::new($sheetXml.NameTable)
  $sheetNs.AddNamespace("s", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")

  $records = @(
    foreach ($row in $sheetXml.SelectNodes("//s:sheetData/s:row[position() > 1]", $sheetNs)) {
      $cells = @{}
      foreach ($cell in $row.SelectNodes("./s:c", $sheetNs)) {
        $valueNode = $cell.SelectSingleNode("./s:v", $sheetNs)
        if (-not $valueNode) { continue }
        $column = ([regex]::Match([string]$cell.r, "^[A-Z]+")).Value
        $value = $valueNode.InnerText
        if ($cell.t -eq "s") { $value = $sharedStrings[[int]$value] }
        $cells[$column] = Clean-Text $value
      }

      $state = $cells["A"]
      if (-not $stateSet.ContainsKey($state)) { continue }

      $sourceValues = @(
        foreach ($column in @("O", "P", "Q", "R", "S", "T", "U", "V", "W")) {
          $candidate = Clean-Text $cells[$column]
          if ($candidate -and $candidate -ne "NA") { $candidate }
        }
      )

      [ordered]@{
        state = $state
        statute = Clean-Text $cells["B"]
        regulation = Clean-Text $cells["C"]
        status = Clean-Text $cells["D"]
        bill = Clean-Text $cells["E"]
        year = Clean-Text $cells["F"]
        sponsor = Clean-Text $cells["G"]
        committee = Clean-Text $cells["H"]
        act = Clean-Text $cells["I"]
        code = Clean-Text $cells["J"]
        driver = Clean-Text $cells["K"]
        liability = Clean-Text $cells["L"]
        commercial = Clean-Text $cells["M"]
        procedure = Clean-Text $cells["N"]
        sources = $sourceValues
      }
    }
  )

  $outputDirectory = Split-Path -Parent $OutputPath
  if ($outputDirectory) { New-Item -ItemType Directory -Force -Path $outputDirectory | Out-Null }
  $json = $records | ConvertTo-Json -Depth 5
  [IO.File]::WriteAllText(
    (Join-Path (Get-Location) $OutputPath),
    $json,
    [Text.UTF8Encoding]::new($false)
  )
  Write-Output "Extracted $($records.Count) state records to $OutputPath"
}
finally {
  $archive.Dispose()
  $stream.Dispose()
}

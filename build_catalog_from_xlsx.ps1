Add-Type -AssemblyName System.IO.Compression.FileSystem

$xlsxPath = Join-Path $PSScriptRoot "materias_talleres.xlsx"
$outPath = Join-Path $PSScriptRoot "materias_catalogo.js"

function Get-ColLetters([string]$cellRef) {
  return ($cellRef -replace '\d', '')
}

function Get-CellValue($cell, $sharedStrings) {
  if ($null -eq $cell) { return "" }

  $cellType = [string]$cell.t
  if ($cellType -eq "inlineStr") {
    return [string]$cell.is.t
  }

  $raw = [string]$cell.v
  if ([string]::IsNullOrWhiteSpace($raw)) { return "" }

  if ($cellType -eq "s") {
    $idx = [int]$raw
    if ($idx -ge 0 -and $idx -lt $sharedStrings.Count) {
      return [string]$sharedStrings[$idx]
    }
  }

  return $raw
}

function Normalize-Header([string]$text) {
  if ($null -eq $text) { return "" }
  $t = $text.Trim().ToLowerInvariant()
  $normalized = $t.Normalize([Text.NormalizationForm]::FormD)
  $sb = New-Object Text.StringBuilder
  foreach ($ch in $normalized.ToCharArray()) {
    $cat = [Globalization.CharUnicodeInfo]::GetUnicodeCategory($ch)
    if ($cat -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
      [void]$sb.Append($ch)
    }
  }

  return $sb.ToString().Replace("ñ", "n")
}

function Find-Column([hashtable]$headerByCol, [string[]]$candidates) {
  foreach ($col in $headerByCol.Keys) {
    $normalized = Normalize-Header([string]$headerByCol[$col])
    foreach ($candidate in $candidates) {
      if ($normalized -eq (Normalize-Header $candidate)) {
        return $col
      }
    }
  }
  return $null
}

function Normalize-Commission([string]$value) {
  $v = [string]$value
  if ([string]::IsNullOrWhiteSpace($v)) { return "" }
  $v = $v.Trim()
  $compact = ($v -replace "\s+", "")
  if ($compact -match "^\d+([\.,]0+)?$") {
    $asDouble = 0.0
    if ([double]::TryParse($compact.Replace(',', '.'), [Globalization.NumberStyles]::Float, [Globalization.CultureInfo]::InvariantCulture, [ref]$asDouble)) {
      return ([int][Math]::Round($asDouble)).ToString()
    }
  }
  return $v
}

if (-not (Test-Path $xlsxPath)) {
  throw "No existe el archivo: $xlsxPath"
}

$zip = [System.IO.Compression.ZipFile]::OpenRead($xlsxPath)

try {
  $sharedStrings = @()
  $sharedEntry = $zip.Entries | Where-Object { $_.FullName -eq "xl/sharedStrings.xml" } | Select-Object -First 1

  if ($sharedEntry) {
    $sr = New-Object System.IO.StreamReader($sharedEntry.Open())
    $sharedText = $sr.ReadToEnd()
    $sr.Close()

    [xml]$sharedXml = $sharedText
    $sharedNs = New-Object System.Xml.XmlNamespaceManager($sharedXml.NameTable)
    $sharedNs.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")

    $siNodes = $sharedXml.SelectNodes("//x:sst/x:si", $sharedNs)
    foreach ($si in $siNodes) {
      $tNodes = $si.SelectNodes(".//x:t", $sharedNs)
      if ($tNodes.Count -gt 0) {
        $sharedStrings += (($tNodes | ForEach-Object { $_."#text" }) -join "")
      } else {
        $sharedStrings += ""
      }
    }
  }

  $sheetEntry = $zip.Entries | Where-Object { $_.FullName -eq "xl/worksheets/sheet1.xml" } | Select-Object -First 1
  if (-not $sheetEntry) {
    throw "No se encontro xl/worksheets/sheet1.xml"
  }

  $srSheet = New-Object System.IO.StreamReader($sheetEntry.Open())
  $sheetText = $srSheet.ReadToEnd()
  $srSheet.Close()

  [xml]$sheetXml = $sheetText
  $ns = New-Object System.Xml.XmlNamespaceManager($sheetXml.NameTable)
  $ns.AddNamespace("x", "http://schemas.openxmlformats.org/spreadsheetml/2006/main")

  $rows = $sheetXml.SelectNodes("//x:worksheet/x:sheetData/x:row", $ns)
  if ($rows.Count -lt 2) {
    throw "La hoja no tiene datos suficientes."
  }

  $headerByCol = @{}
  $headerRow = $rows[0]
  foreach ($cell in $headerRow.SelectNodes("x:c", $ns)) {
    $col = Get-ColLetters([string]$cell.r)
    $headerByCol[$col] = (Get-CellValue $cell $sharedStrings).Trim()
  }

  $colProfesor = Find-Column $headerByCol @("apellidos y nombres", "profesor", "docente")
  $colMateria = Find-Column $headerByCol @("cargos / asignaturas", "asignatura", "materia", "instrumento/materia")
  $colAnio = Find-Column $headerByCol @("año", "anio")
  $colComision = Find-Column $headerByCol @("div.", "div", "division", "división", "comision")
  $colAlumnos = Find-Column $headerByCol @("alumnos", "cantidad de alumnos", "cupo")

  # Fallback para el formato fijo de materias_talleres.xlsx (A: Año, C: Materia, D: Profesor, E: División, F: Alumnos)
  if (-not $colAnio -and $headerByCol.ContainsKey("A")) { $colAnio = "A" }
  if (-not $colMateria -and $headerByCol.ContainsKey("C")) { $colMateria = "C" }
  if (-not $colProfesor -and $headerByCol.ContainsKey("D")) { $colProfesor = "D" }
  if (-not $colComision -and $headerByCol.ContainsKey("E")) { $colComision = "E" }
  if (-not $colAlumnos -and $headerByCol.ContainsKey("F")) { $colAlumnos = "F" }

  if (-not $colProfesor -or -not $colMateria -or -not $colAnio) {
    $detected = ($headerByCol.GetEnumerator() | ForEach-Object { "[$($_.Key)] $($_.Value)" }) -join ", "
    throw "No se pudieron mapear columnas requeridas. Encabezados detectados: $detected"
  }

  $catalog = [ordered]@{}

  for ($rowIndex = 1; $rowIndex -lt $rows.Count; $rowIndex++) {
    $row = $rows[$rowIndex]
    $cellByCol = @{}

    foreach ($cell in $row.SelectNodes("x:c", $ns)) {
      $cellByCol[(Get-ColLetters([string]$cell.r))] = $cell
    }

    $profesor = if ($cellByCol.ContainsKey($colProfesor)) { (Get-CellValue $cellByCol[$colProfesor] $sharedStrings).Trim() } else { "" }
    $materia = if ($cellByCol.ContainsKey($colMateria)) { (Get-CellValue $cellByCol[$colMateria] $sharedStrings).Trim() } else { "" }
    $anioRaw = if ($cellByCol.ContainsKey($colAnio)) { (Get-CellValue $cellByCol[$colAnio] $sharedStrings).Trim() } else { "" }
    $comision = if ($colComision -and $cellByCol.ContainsKey($colComision)) { (Get-CellValue $cellByCol[$colComision] $sharedStrings).Trim() } else { "" }
    $alumnosRaw = if ($colAlumnos -and $cellByCol.ContainsKey($colAlumnos)) { (Get-CellValue $cellByCol[$colAlumnos] $sharedStrings).Trim() } else { "" }

    if ([string]::IsNullOrWhiteSpace($profesor) -or [string]::IsNullOrWhiteSpace($materia)) {
      continue
    }

    $anio = $null
    if ($anioRaw -match "(\d+)") {
      $anio = [int]$Matches[1]
    }
    if ($null -eq $anio) {
      continue
    }

    $cupo = $null
    if (-not [string]::IsNullOrWhiteSpace($alumnosRaw)) {
      $n = 0
      if ([int]::TryParse($alumnosRaw, [ref]$n)) {
        $cupo = $n
      }
    }

    if (-not $catalog.Contains($materia)) {
      $catalog[$materia] = [ordered]@{
        type = $(if ($cupo -eq $null) { "group" } else { "individual" })
        description = ""
        years = @()
        allow_over = $true
        professors = [ordered]@{}
        commissions = [ordered]@{}
      }
    }

    $entry = $catalog[$materia]

    if ($entry.years -notcontains $anio) {
      $entry.years += $anio
      $entry.years = @($entry.years | Sort-Object)
    }

    if (-not $entry.professors.Contains($profesor)) {
      $entry.professors[$profesor] = [ordered]@{}
    }

    $comision = Normalize-Commission $comision

    if (-not [string]::IsNullOrWhiteSpace($comision)) {
      if (-not $entry.commissions.Contains($comision)) {
        $entry.commissions[$comision] = [ordered]@{
          capacity = $cupo
          professor = $profesor
        }
      }
    }

    if ($entry.type -ne "group" -and $cupo -eq $null) {
      $entry.type = "group"
    }
  }

  $json = $catalog | ConvertTo-Json -Depth 10
  $content = "window.MATERIAS_CATALOGO = $json`r`n"
  Set-Content -Path $outPath -Value $content -Encoding UTF8

  Write-Output ("Catalogo generado. Materias: " + $catalog.Keys.Count)
}
finally {
  $zip.Dispose()
}

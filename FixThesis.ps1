# =====================================================================
#  FixThesis.ps1  (focused version)
#  - Figure/Table caption format fix (SEQ field kept intact)
#  - Equation numbering: re-numbered from scratch based on which
#    chapter/section each equation actually falls under, strictly
#    sequential within that section (no gaps, no jumps).
#  Headings are NOT touched here (handled separately via VBA).
#  In-text reference review is NOT included in this version.
# =====================================================================

param(
    [string]$DocPath = ""
)

$ErrorActionPreference = "Stop"

# ---------- 0) Find the file ----------
if ([string]::IsNullOrWhiteSpace($DocPath)) {
    $candidates = Get-ChildItem -Path $PSScriptRoot -Filter *.docx | Where-Object { $_.Name -notmatch "^~\$" }
    if ($candidates.Count -eq 1) {
        $DocPath = $candidates[0].FullName
    } elseif ($candidates.Count -gt 1) {
        Write-Host "Multiple docx files found in this folder:"
        for ($i=0; $i -lt $candidates.Count; $i++) { Write-Host "  [$i] $($candidates[$i].Name)" }
        $sel = Read-Host "Enter the number of the file to use"
        $DocPath = $candidates[[int]$sel].FullName
    } else {
        $DocPath = Read-Host "Enter the full path to the thesis .docx file"
    }
}

if (-not (Test-Path $DocPath)) {
    Write-Host "File not found: $DocPath" -ForegroundColor Red
    exit 1
}

# ---------- 1) Backup ----------
$backupPath = [System.IO.Path]::Combine(
    [System.IO.Path]::GetDirectoryName($DocPath),
    [System.IO.Path]::GetFileNameWithoutExtension($DocPath) + "_backup_" + (Get-Date -Format "yyyyMMdd_HHmmss") + ".docx"
)
Copy-Item -Path $DocPath -Destination $backupPath
Write-Host "Backup created:" -ForegroundColor Green
Write-Host "  $backupPath`n"

Read-Host "Press Enter to start"

# ---------- helper: convert Persian/Arabic-Indic digits to Latin ----------
function ConvertDigitsToLatin([string]$s) {
    $persian = @('۰','۱','۲','۳','۴','۵','۶','۷','۸','۹')
    $arabic  = @('٠','١','٢','٣','٤','٥','٦','٧','٨','٩')
    $out = $s
    for ($d = 0; $d -le 9; $d++) {
        $out = $out.Replace($persian[$d], [string]$d)
        $out = $out.Replace($arabic[$d], [string]$d)
    }
    return $out
}

# ---------- 2) Open Word ----------
Write-Host "`nOpening Word ..." -ForegroundColor Cyan
$word = New-Object -ComObject Word.Application
$word.Visible = $true
$doc = $word.Documents.Open($DocPath)
$doc.TrackRevisions = $false

$wdFieldSequence = 12

# =====================================================================
# PART 1: Figure / Table caption format
# Insert-only edits (never delete/replace a range touching the field
# boundary), so the SEQ field (and therefore Update Table) stays intact.
# =====================================================================
Write-Host "`n[1] Fixing figure/table caption format ..." -ForegroundColor Cyan

$fixedCaptions = 0
$skippedTables = 0
$fieldsSnapshot = @($doc.Fields)
foreach ($field in $fieldsSnapshot) {
    try {
        if ($field.Type -ne $wdFieldSequence) { continue }
        $code = $field.Code.Text
        if ($code -notmatch "SEQ\s+(Figure|Table)") { continue }
        $kind = $Matches[1]

        $codeStart = $field.Code.Start
        $paraStart = $field.Code.Paragraphs.Item(1).Range.Start
        $prefixRange = $doc.Range($paraStart, $codeStart)
        $prefixText = $prefixRange.Text

        if ($prefixText -match '^(?<label>.*?\s)(?<mark>[\u200e\u200f]*)(?<num>[0-9]+)(?<sep>[.\u200c]+)\s*$') {
            if ($prefixText -notmatch '^\(') {
                $labelLen = $Matches['label'].Length
                $markLen  = $Matches['mark'].Length
                $numLen   = $Matches['num'].Length
                $sepLen   = $Matches['sep'].Length

                $numStart = $paraStart + $labelLen + $markLen
                $sepStart = $numStart + $numLen
                $sepEnd   = $sepStart + $sepLen

                # Edit order matters: rightmost edit first, so earlier
                # positions we already computed are never shifted.
                $resultEnd = $field.Result.End
                $doc.Range($resultEnd, $resultEnd).InsertAfter('):')
                $doc.Range($sepStart, $sepEnd).Text = '-'
                $doc.Range($numStart, $numStart).InsertBefore('(')

                $fixedCaptions++
            }
        } elseif ($prefixText -match '^(?<label>.*?)\((?<mark>[\u200e\u200f]*)(?<num>[0-9]+)(?<sep>[.\u200c\-]+)\s*$') {
            # already partially converted by an older/previous run: has "("
            # and possibly "-" already, just needs the closing "):" appended.
            $resultText = $field.Result.Text
            if ($resultText -notmatch '\)' -and $prefixText -notmatch '\)') {
                $resultEnd = $field.Result.End
                $doc.Range($resultEnd, $resultEnd).InsertAfter('):')
                $fixedCaptions++
            }
        } elseif ($kind -eq "Table") {
            $skippedTables++
            Write-Host ("  [DIAGNOSTIC] Table caption did not match, raw prefix text: [" + $prefixText + "]") -ForegroundColor DarkYellow
        }
    } catch {
        Write-Host "  Skipped one caption: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}
Write-Host "  $fixedCaptions caption(s) fixed." -ForegroundColor Green
if ($skippedTables -gt 0) {
    Write-Host "  $skippedTables table caption(s) skipped - see [DIAGNOSTIC] lines above." -ForegroundColor DarkYellow
}

# =====================================================================
# PART 2: Equation numbering
# Equations are plain text (no automatic field), so we fully re-number
# them: walk the document in order, track the current chapter from the
# nearest preceding Heading 1, and assign a strictly sequential counter
# per chapter (same two-part "(chapter-counter):" shape as figures/tables).
# =====================================================================
Write-Host "`n[2] Re-numbering equations by chapter ..." -ForegroundColor Cyan

$chapterNum = "0"
$counters = @{}
$fixedEquations = 0

foreach ($p in $doc.Paragraphs) {
    $lvl = $p.OutlineLevel
    if ($lvl -eq 1 -and $p.Range.ListFormat.ListType -ne 0) {
        $listStr = ConvertDigitsToLatin $p.Range.ListFormat.ListString
        $nums = [regex]::Matches($listStr, '[0-9]+') | ForEach-Object { $_.Value }
        if ($nums.Count -ge 1) { $chapterNum = $nums[0] }
        continue
    }

    $styleName = ""
    try { $styleName = $p.Style.NameLocal } catch { $styleName = "" }
    if ($styleName -notmatch "Equa") { continue }

    $pText = $p.Range.Text
    $m = [regex]::Match($pText, '\([0-9]{1,2}[.\-][0-9]{1,2}(?:[.\-][0-9]{1,2})?\)\:?')
    if (-not $m.Success) { continue }

    $key = "$chapterNum"
    if (-not $counters.ContainsKey($key)) { $counters[$key] = 0 }
    $counters[$key] = $counters[$key] + 1

    $newNum = "($chapterNum-$($counters[$key])):"

    $absStart = $p.Range.Start + $m.Index
    $absEnd = $absStart + $m.Length
    try {
        # Delete then Insert (same style as the caption fix), rather than
        # a direct Text= replacement, so bidi/ordering behaves the same
        # way it does for figures/tables.
        $targetRange = $doc.Range($absStart, $absEnd)
        [void]$targetRange.Delete()
        [void]$doc.Range($absStart, $absStart).InsertBefore($newNum)

        # A freshly-inserted run does not automatically inherit the RTL
        # (Bidi) property the surrounding original text had, which is
        # what caused the reversed/garbled display. Force it explicitly.
        $newRange = $doc.Range($absStart, $absStart + $newNum.Length)
        $newRange.Bidi = $true
        $newRange.Font.Name = "Times New Roman"

        $fixedEquations++
    } catch {
        Write-Host "  Skipped one equation: $($_.Exception.Message)" -ForegroundColor Yellow
    }
}
Write-Host "  $fixedEquations equation(s) re-numbered." -ForegroundColor Green

# ---------- Save ----------
Write-Host "`nSaving ..." -ForegroundColor Cyan
$doc.Save()

Write-Host "`nDone." -ForegroundColor Green
Write-Host "Backup is here if anything needs reverting:"
Write-Host "  $backupPath"
Write-Host "`nReminder: press Ctrl+A then F9 once in Word to refresh the List of Figures/Tables."
Read-Host "`nPress Enter to close"

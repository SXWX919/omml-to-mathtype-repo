param(
    [Parameter(Mandatory = $true)][string]$DocName,
    [string]$ButtonName = '切换 TeX',
    [int]$DelayMs = 1700,
    [switch]$NoSave
)
# Convert LaTeX marker text (\[...\] / $...$) inside an open Word document into
# MathType OLE equations by invoking MathType's "Toggle TeX" ribbon command via UIA.
# No mouse movement. Word must have been started normally (not via COM /Automation).

$ErrorActionPreference = 'Continue'
Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes

$app = $null
try { $app = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application') }
catch { "cannot attach to a running Word instance: $($_.Exception.Message)"; exit 1 }
$app.DisplayAlerts = 0

$doc = $null
foreach ($d in $app.Documents) { if ($d.Name -eq $DocName) { $doc = $d } }
if (-not $doc) { "document not open in Word: $DocName"; exit 1 }
"doc: $($doc.Name)"

$AE = [System.Windows.Automation.AutomationElement]
$TS = [System.Windows.Automation.TreeScope]
$root = $AE::RootElement
$all = $root.FindAll($TS::Children, [System.Windows.Automation.Condition]::TrueCondition)
$win = $null
foreach ($w in $all) {
    if ($w.Current.ClassName -eq 'OpusApp' -and $w.Current.Name -like "*$DocName*") { $win = $w }
}
if (-not $win) { "no Word window for $DocName (is the document visible?)"; exit 1 }

$els = $win.FindAll($TS::Descendants, [System.Windows.Automation.Condition]::TrueCondition)
$tab = $null
foreach ($e in $els) {
    if ($e.Current.ControlType.ProgrammaticName -eq 'ControlType.TabItem' -and $e.Current.Name -eq 'MathType') { $tab = $e; break }
}
if (-not $tab) { "MathType ribbon tab not found - start Word normally, not via COM /Automation"; exit 1 }
$tab.GetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern).Select()
Start-Sleep -Milliseconds 1200

$els = $win.FindAll($TS::Descendants, [System.Windows.Automation.Condition]::TrueCondition)
$btn = $null
foreach ($e in $els) {
    if ($e.Current.ControlType.ProgrammaticName -eq 'ControlType.Button' -and $e.Current.Name -eq $ButtonName) { $btn = $e; break }
}
if (-not $btn) { "button not found: $ButtonName"; exit 1 }
$inv = $btn.GetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern)

$pattern = '\\\[.*?\\\]|\$[^$]*?\$'
$rx = New-Object System.Text.RegularExpressions.Regex($pattern, [System.Text.RegularExpressions.RegexOptions]::Singleline)

$done = 0; $fail = 0; $failList = @()
$paras = $doc.Paragraphs
for ($pi = $paras.Count; $pi -ge 1; $pi--) {
    $p = $paras.Item($pi)
    $txt = $p.Range.Text
    $ms = $rx.Matches($txt)
    if ($ms.Count -eq 0) { continue }
    $base = $p.Range.Start
    for ($k = $ms.Count - 1; $k -ge 0; $k--) {
        $m = $ms[$k]
        $marker = $m.Value
        $before = $doc.InlineShapes.Count
        $r = $null
        try {
            $r = $doc.Range($base + $m.Index, $base + $m.Index + $m.Length)
            $r.Select()
            if ($r.Text -ne $marker) {
                # character offsets drift in paragraphs containing fields / inline objects
                $fr = $p.Range
                $f = $fr.Find
                $f.ClearFormatting()
                $f.Text = $marker
                $f.Forward = $true
                $f.Wrap = 0
                $f.MatchWildcards = $false
                if ($f.Execute()) { $fr.Select(); $r = $fr } else { $r = $null }
            }
            if ($null -eq $r) { throw 'marker not selectable' }
            Start-Sleep -Milliseconds 200
            $inv.Invoke()
            Start-Sleep -Milliseconds $DelayMs
        }
        catch {
            $fail++; $failList += $marker
            "ERR  : " + $marker.Substring(0, [Math]::Min(70, $marker.Length)) + " | " + $_.Exception.Message
            continue
        }
        $after = $doc.InlineShapes.Count
        if ($after -gt $before) {
            $done++
            "OK   $done : " + $marker.Substring(0, [Math]::Min(50, $marker.Length))
        }
        else {
            $fail++; $failList += $marker
            "FAIL : " + $marker.Substring(0, [Math]::Min(70, $marker.Length))
        }
    }
}
if (-not $NoSave) { $doc.Save() }
"SUMMARY done=$done fail=$fail inlineshapes=$($doc.InlineShapes.Count)"
if ($fail -gt 0) {
    "FAILED markers (rerun after fixing: bare caps -> \mathrm{...}, or convert them manually):"
    foreach ($t in $failList) { "  " + $t.Substring(0, [Math]::Min(90, $t.Length)) }
}

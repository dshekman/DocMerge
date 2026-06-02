# DocMerge Launcher
# Downloads the latest merge_tool.py from GitHub and runs it.
# Requirements: Python must be installed (python.org)

$url     = "https://raw.githubusercontent.com/dshekman/DocMerge/main/merge_tool.py"
$tmpFile = "$env:TEMP\merge_tool.py"

# ── Download ──────────────────────────────────────────────────────────────────
Write-Host "Downloading latest DocMerge..." -ForegroundColor Cyan
try {
    Invoke-WebRequest -Uri $url -OutFile $tmpFile -UseBasicParsing
} catch {
    Write-Host "ERROR: Could not download merge_tool.py" -ForegroundColor Red
    Write-Host $_.Exception.Message
    pause; exit 1
}

# ── Find Python ───────────────────────────────────────────────────────────────
$python = $null
foreach ($cmd in @("python", "python3", "py")) {
    if (Get-Command $cmd -ErrorAction SilentlyContinue) {
        $python = $cmd; break
    }
}
if (-not $python) {
    Write-Host "ERROR: Python not found. Install it from https://python.org" -ForegroundColor Red
    pause; exit 1
}

# ── Install python-docx (ignore errors — may already be present) ──────────────
Write-Host "Checking dependencies..." -ForegroundColor Cyan
try {
    $out = & $python -m pip install python-docx --quiet 2>&1
} catch { }

# ── Launch ────────────────────────────────────────────────────────────────────
Write-Host "Launching DocMerge..." -ForegroundColor Green
& $python $tmpFile

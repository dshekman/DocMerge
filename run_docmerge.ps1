# DocMerge Launcher
# Downloads the latest merge_tool.py from GitHub and runs it.
# Requirements: Python must be installed (python.org)

$ErrorActionPreference = "Stop"

$url     = "https://raw.githubusercontent.com/dshekman/DocMerge/main/merge_tool.py"
$tmpFile = "$env:TEMP\merge_tool.py"

Write-Host "Downloading latest DocMerge..." -ForegroundColor Cyan

try {
    Invoke-WebRequest -Uri $url -OutFile $tmpFile -UseBasicParsing
} catch {
    Write-Host "ERROR: Could not download merge_tool.py" -ForegroundColor Red
    Write-Host $_.Exception.Message
    pause
    exit 1
}

# Check Python is available
$python = $null
foreach ($cmd in @("python", "python3", "py")) {
    if (Get-Command $cmd -ErrorAction SilentlyContinue) {
        $python = $cmd
        break
    }
}

if (-not $python) {
    Write-Host "ERROR: Python not found. Install it from https://python.org" -ForegroundColor Red
    pause
    exit 1
}

# Install python-docx if missing
Write-Host "Checking dependencies..." -ForegroundColor Cyan
$checkDocx = & $python -c "import docx; print('ok')" 2>&1
if ($checkDocx -notmatch "ok") {
    Write-Host "Installing python-docx..." -ForegroundColor Yellow
    & $python -m pip install python-docx --quiet
}

Write-Host "Launching DocMerge..." -ForegroundColor Green
& $python $tmpFile

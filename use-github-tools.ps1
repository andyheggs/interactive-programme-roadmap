$ProjectRoot = $PSScriptRoot
$GitPath = Join-Path $ProjectRoot ".tools\MinGit\cmd"
$GhPath = Join-Path $ProjectRoot ".tools\gh\bin"

$env:Path = "$GitPath;$GhPath;$env:Path"

Write-Host "Local Git/GitHub tools enabled for this PowerShell session."
git --version
gh --version

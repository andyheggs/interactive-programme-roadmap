$ProjectRoot = $PSScriptRoot
$GitPath = Join-Path $ProjectRoot ".tools\MinGit\cmd"
$GhPath = Join-Path $ProjectRoot ".tools\gh\bin"
$NodePath = Join-Path $ProjectRoot ".tools\node"

$env:Path = "$GitPath;$GhPath;$NodePath;$env:Path"

Write-Host "Local Git/GitHub/Node tools enabled for this PowerShell session."
git --version
gh --version
node --version
npm --version

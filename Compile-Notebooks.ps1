$folder = ".\Monthly Run\Cash\Compiled Script (Beta)"
$output = "$folder\Compiled_All_Notebooks.ipynb"

$files = Get-ChildItem $folder -Filter *.ipynb | Sort-Object Name

$base = Get-Content $files[0].FullName -Raw | ConvertFrom-Json

foreach ($f in $files[1..($files.Count - 1)]) {
 $nb = Get-Content $f.FullName -Raw | ConvertFrom-Json
 $base.cells += $nb.cells
}

$base | ConvertTo-Json -Depth 100 | Set-Content $output -Encoding utf8

Write-Host "Done. New notebook created at: $output"

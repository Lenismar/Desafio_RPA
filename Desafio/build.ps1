$exclude = @("venv", "Desafio.zip")
$files = Get-ChildItem -Path . -Exclude $exclude
Compress-Archive -Path $files -DestinationPath "Desafio.zip" -Force
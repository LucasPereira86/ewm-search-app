Add-Type -AssemblyName System.Drawing

# Convert icon-192
$img192 = [System.Drawing.Image]::FromFile("C:\Users\Cliente\.gemini\antigravity\scratch\excel-search-app\icon-192.png")
$img192.Save("C:\Users\Cliente\.gemini\antigravity\scratch\excel-search-app\icon-192-converted.png", [System.Drawing.Imaging.ImageFormat]::Png)
$img192.Dispose()

# Convert icon-512
$img512 = [System.Drawing.Image]::FromFile("C:\Users\Cliente\.gemini\antigravity\scratch\excel-search-app\icon-512.png")
$img512.Save("C:\Users\Cliente\.gemini\antigravity\scratch\excel-search-app\icon-512-converted.png", [System.Drawing.Imaging.ImageFormat]::Png)
$img512.Dispose()

# Replace original files
Remove-Item "C:\Users\Cliente\.gemini\antigravity\scratch\excel-search-app\icon-192.png" -Force
Remove-Item "C:\Users\Cliente\.gemini\antigravity\scratch\excel-search-app\icon-512.png" -Force
Rename-Item "C:\Users\Cliente\.gemini\antigravity\scratch\excel-search-app\icon-192-converted.png" "icon-192.png"
Rename-Item "C:\Users\Cliente\.gemini\antigravity\scratch\excel-search-app\icon-512-converted.png" "icon-512.png"

Write-Host "Icons converted to PNG successfully!"

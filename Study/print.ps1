Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Captura o nome de usuário do computador
$usuario = $env:USERNAME


Write-Host $usuario

# Criar diretório na rede para salvar os prints
$diretorioRede = "S:\PS\qmc\Dados_qmc\60_TI\62_Controles\05_5s Eletronico\$usuario"


Write-Host $diretorioRede



# Verificar se o diretório existe, se não, criar
if (-not (Test-Path $diretorioRede -PathType Container)) {
    New-Item -ItemType Directory -Path $diretorioRede -ErrorAction SilentlyContinue
}

# Função para tirar um print de uma janela específica
function Capture-Screenshot {
    param(
        [string]$windowTitle,
        [string]$outputFilePath
    )

    $window = Get-Process | Where-Object { $_.MainWindowTitle -eq $windowTitle } | Select-Object -First 1
    if ($window) {
        $windowHandle = $window.MainWindowHandle
        $windowBounds = [System.Windows.Forms.Screen]::FromHandle($windowHandle).Bounds
        $bitmap = New-Object System.Drawing.Bitmap($windowBounds.Width, $windowBounds.Height)
        $graphics = [System.Drawing.Graphics]::FromImage($bitmap)
        $graphics.CopyFromScreen($windowBounds.Location, [System.Drawing.Point]::Empty, $windowBounds.Size)
        $bitmap.Save($outputFilePath, [System.Drawing.Imaging.ImageFormat]::Png)
    } else {
        Write-Host "Janela '$windowTitle' não encontrada."
    }
}

# Abrir as janelas
Start-Process "explorer.exe" "$env:USERPROFILE\Documents"
Start-Process "explorer.exe" "shell:RecycleBinFolder"
Start-Process "explorer.exe" "$env:USERPROFILE\Pictures"
Start-Process "explorer.exe" "$env:USERPROFILE\Downloads"

# Aguardar um momento para as janelas abrirem completamente
Start-Sleep -Seconds 5

# Tirar prints das janelas e salvar na pasta na rede
Capture-Screenshot "Documents" "$diretorioRede\Documents.png"
Capture-Screenshot "Pictures" "$diretorioRede\Pictures.png"
Capture-Screenshot "Downloads" "$diretorioRede\Downloads.png"

# Fechar as janelas
Stop-Process -Name "explorer" -ErrorAction SilentlyContinue

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Captura o nome de usuário do computador
$usuario = $env:USERNAME

# Criar diretório na rede para salvar os prints
$diretorioRede = "\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\$usuario"

# Verificar se o diretório existe, se não, criar
if (-not (Test-Path $diretorioRede -PathType Container)) {
    New-Item -ItemType Directory -Path $diretorioRede -ErrorAction SilentlyContinue
}

# Função para tirar um print de uma janela específica
function Capture-DualMonitorScreenshot {
    param(
        [string]$outputFilePath
    )

    $screenBounds = [System.Windows.Forms.SystemInformation]::VirtualScreen

    $bitmap = New-Object System.Drawing.Bitmap($screenBounds.Width, $screenBounds.Height)
    $graphics = [System.Drawing.Graphics]::FromImage($bitmap)

    $graphics.CopyFromScreen($screenBounds.Location, [System.Drawing.Point]::Empty, $screenBounds.Size)

    $bitmap.Save($outputFilePath, [System.Drawing.Imaging.ImageFormat]::Png)
}

#minimizar todos os programas abertos

Capture-DualMonitorScreenshot "$diretorioRede\Desktop.png"

# Mapeia os caminhos das pastas especiais e seus títulos correspondentes
$specialFolders = @{
    'Desktop' = [System.Environment]::GetFolderPath('Desktop'), 'Desktop'
    'Documents' = [System.Environment]::GetFolderPath('MyDocuments'), 'Documents'
    'Pictures' = [System.Environment]::GetFolderPath('MyPictures'), 'Pictures'
    'Downloads' = "$env:USERPROFILE\Downloads", 'Downloads'
    'Recycle Bin' = "shell:RecycleBinFolder", 'Recycle Bin'
}

# Abre as janelas e captura os prints
foreach ($folder in $specialFolders.GetEnumerator()) {
    $folderPath = $folder.Value[0]

    if ($folderPath -ne $null) {
        Start-Process "explorer.exe" $folderPath
        Start-Sleep -Seconds 4
        Capture-DualMonitorScreenshot "$diretorioRede\$($folder.Key).png"
    } else {
        Write-Host "Caminho da pasta '$($folder.Key)' não encontrado."
    }
}

# Fechar as janelas
Stop-Process -Name "explorer" -ErrorAction SilentlyContinue

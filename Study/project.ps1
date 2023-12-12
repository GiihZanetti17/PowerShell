Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Captura o nome de usuário do computador
$usuario = $env:USERNAME

# Criar diretório na rede para salvar os prints
$diretorioRede = "S:\PS\qmc\Dados_qmc\60_TI\62_Controles\05_5s Eletronico\$usuario"

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

# Minimizar todas as janelas abertas
$shell = New-Object -ComObject Shell.Application
$shell.MinimizeAll()

Start-Sleep -Seconds 2  # Tempo adicional para as janelas minimizarem completamente

# Capturar print das duas telas minimizadas
Capture-DualMonitorScreenshot "$diretorioRede\DualScreen_Minimized.png"

# Aumentar as janelas antes de capturar os prints das pastas específicas
$wshell = New-Object -ComObject wscript.shell

# Mapeia os caminhos das pastas especiais e seus títulos correspondentes
$specialFolders = @{
    'Desktop' = [System.Environment]::GetFolderPath('Desktop')
    'Documents' = [System.Environment]::GetFolderPath('MyDocuments')
    'Pictures' = [System.Environment]::GetFolderPath('MyPictures')
    'Downloads' = "$env:USERPROFILE\Downloads"
    'Recycle Bin' = "shell:RecycleBinFolder"
}

# Abre as janelas e captura os prints
foreach ($folder in $specialFolders.GetEnumerator()) {
    $folderName = $folder.Key
    $folderPath = $folder.Value

    if ($folderPath -ne $null) {
        # Abre a janela e espera um pouco antes de capturar o print
        Start-Process "explorer.exe" $folderPath
        Start-Sleep -Seconds 2
        
        # Aumenta a janela antes de capturar o print
        $wshell.SendKeys("% x")  # Alt + Espaço (abre o menu da janela)
        Start-Sleep -Milliseconds 500
        $wshell.SendKeys("x")    # Maximize a janela

        Start-Sleep -Seconds 2  # Tempo para a janela maximizar completamente

        Capture-DualMonitorScreenshot "$diretorioRede\$folderName.png"
    } else {
        Write-Host "Caminho da pasta '$folderName' não encontrado."
    }
}

# Fechar as janelas do Explorer
Stop-Process -Name "explorer" -ErrorAction SilentlyContinue

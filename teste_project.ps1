Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Captura o nome de usuário do computador
$usuario = $env:USERNAME

# Criar diretório na rede para salvar os prints
$diretorioRede = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx\$usuario"



# Perguntas
$questions = @(
    "Coloque as principais pastas que você utiliza.",
    "Pastas estão padronizadas?",
    "Contém arquivos duplicados?",
    "Contém arquivos com mais de 5 anos?",
    "Contém arquivos que não são utilizados?",
    "Informe a 1º pasta para você organizar e limpar.",
    "Informe a 2º pasta para você organizar e limpar.",
    "Está com o histórico do navegador limpo?",
    "Pontos de Melhoria",
    "Computador/Notebook está travado com o cadeado?"
)

# Array para armazenar respostas
$answers = @()

# Loop para exibir as perguntas e receber as respostas
foreach ($question in $questions) {
    $answer = Read-Host -Prompt $question
    $answers += $answer
}

# Salvar as respostas em um arquivo de texto
$answersPath = Join-Path -Path $diretorioRede -ChildPath "$usuario-respostas.txt"
$answers | Out-File -FilePath $answersPath

# Criar um arquivo de log para verificar o processo
$logPath = Join-Path -Path $diretorioRede -ChildPath "$usuario-log.txt"
"Script executado em: $(Get-Date)" | Out-File -FilePath $logPath -Append



# Pergunta se a pessoa está pronta para a auditoria
$readyForAudit = Read-Host -Prompt "Você está pronto para a auditoria? Responda 'sim' para continuar."

# Verifica a resposta e continua se for 'sim'
if ($readyForAudit -eq 'sim') {
   

   Start-Sleep -Seconds 4  # Tempo adicional para as janelas minimizarem completamente


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

## Abre as janelas e captura os prints
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


    # Exibir mensagem de agradecimento
    [System.Windows.Forms.MessageBox]::Show("Obrigado por completar a auditoria! Os prints foram capturados com sucesso.", "Auditoria Concluída", [System.Windows.Forms.MessageBoxButtons]::OK, [System.Windows.Forms.MessageBoxIcon]::Information)
} else {
    Write-Host "Auditoria cancelada. Os prints não foram capturados."
}

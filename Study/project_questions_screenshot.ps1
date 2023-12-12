Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Captura o nome de usuário do computador
$usuario = $env:USERNAME

# Criar diretório na rede para salvar os prints
$diretorioRede = "S:\PS\qmc\Dados_qmc\60_TI\62_Controles\05_5s Eletronico\$usuario"

# Função para criar a janela de input
function Show-InputBox($message) {
    Add-Type -AssemblyName System.Windows.Forms
    $form = New-Object System.Windows.Forms.Form
    $form.Text = "Question"
    $form.Size = New-Object System.Drawing.Size(300, 150)
    $form.StartPosition = "CenterScreen"
    $label = New-Object System.Windows.Forms.Label
    $label.Location = New-Object System.Drawing.Point(10, 20)
    $label.Size = New-Object System.Drawing.Size(280, 20)
    $label.Text = $message
    $form.Controls.Add($label)
    $textBox = New-Object System.Windows.Forms.TextBox
    $textBox.Location = New-Object System.Drawing.Point(10, 50)
    $textBox.Size = New-Object System.Drawing.Size(260, 20)
    $form.Controls.Add($textBox)
    $button = New-Object System.Windows.Forms.Button
    $button.Location = New-Object System.Drawing.Point(100, 90)
    $button.Size = New-Object System.Drawing.Size(75, 23)
    $button.Text = "OK"
    $button.Add_Click({ $form.Close() })
    $form.Controls.Add($button)
    $form.Add_Shown({ $textBox.Select() })
    [void]$form.ShowDialog()
    return $textBox.Text
}

# Perguntas
$questions = @(
    "Coloque as principais pastas que você utiliza.",
    "Pastas estão padronizadas?",
    "Contém arquivos duplicados?",
    "Contém arquivos com mais de 5 anos?",
    "Contém arquivos que não são utilizados?",
    "Selecione duas pastas para você organizar e limpar.",
    "Está com o histórico do navegador limpo?",
    "Computador/Notebook está travado com o cadeado?"
)

# Array para armazenar respostas
$answers = @()

# Loop para exibir as perguntas e receber as respostas
foreach ($question in $questions) {
    $answer = Show-InputBox -message $question
    $answers += $answer
}

# Criar um objeto Excel
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$workbook = $excel.Workbooks.Add()
$sheet = $workbook.Sheets.Item(1)

# Escrever as perguntas e respostas no Excel
for ($i = 0; $i -lt $questions.Count; $i++) {
    $sheet.Cells.Item($i + 1, 1) = $questions[$i]
    $sheet.Cells.Item($i + 1, 2) = $answers[$i]
}

# Salvar o arquivo Excel na pasta da rede com o usuário
$excelFileName = "$usuario-respostas.xlsx"
$savePath = Join-Path -Path $diretorioRede -ChildPath $excelFileName
$workbook.SaveAs($savePath)
$workbook.Close()
$excel.Quit()


# Limpar objetos Excel da memória
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()


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

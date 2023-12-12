Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Captura o nome de usuário do computador
$usuario = $env:USERNAME

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
$savePath = "S:\PS\qmc\Dados_qmc\60_TI\62_Controles\05_5s Eletronico\$env:USERNAME-respostas.xlsx"
$workbook.SaveAs($savePath)
$workbook.Close()
$excel.Quit()

# Limpar objetos Excel da memória
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

Write-Host "Arquivo Excel criado e salvo em: $savePath"

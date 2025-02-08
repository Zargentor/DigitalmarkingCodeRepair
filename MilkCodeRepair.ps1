# Скрипт для обработки кодов маркировки из TXT файла
# Основные этапы:
# 1. Замена специальных символов
# 2. Удаление лишних кавычек
# 3. Экспорт в TXT и CSV

# Подключение GUI-библиотеки для работы с диалоговыми окнами
Add-Type -AssemblyName System.Windows.Forms

# Диалог выбора исходного файла
# --------------------------------------------------
$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$openFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"  # Фильтр по умолчанию
$result = $openFileDialog.ShowDialog()

# Проверка выбора файла
if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
    Write-Output "Файл не выбран."
    exit
}

# Парсинг пути файла
# --------------------------------------------------
$filePath = $openFileDialog.FileName
$folderPath = [System.IO.Path]::GetDirectoryName($filePath)  # Директория файла
$fileName = [System.IO.Path]::GetFileNameWithoutExtension($filePath)  # Имя без расширения
$fileExtension = [System.IO.Path]::GetExtension($filePath)  # Исходное расширение

# Чтение исходных данных
# --------------------------------------------------
$Codes = Get-Content -Path $filePath  # Загрузка необработанных кодов

# Инициализация временных хранилищ
$tmpMasFirst = @()  # После замены \u001D
$tmpMasSecond = @()  # После замены экранированных кавычек
$tmpMasThird = @()  # После удаления обрамляющих кавычек

# Конфигурация замены символов
# --------------------------------------------------
$CharChangeTo29 = "\u001D"  # Unicode символ Group Separator
$CharBackSlashWhoNeedChange = '\"'  # Экранированные кавычки
$CharBackSlashChangable = '"'       # Обычные кавычки

# Основная обработка данных
# --------------------------------------------------
# Замена Unicode-символа \u001D на реальный символ GS (ASCII 29)
foreach ($Code in $Codes) {
    $tmpMasFirst += $Code -replace [regex]::Escape($CharChangeTo29), [char]29
}

# Экспорт первого этапа обработки
$tmpMasFirst | Out-File -FilePath "$folderPath\$($fileName)_CHG29.txt" -Force

# Замена экранированных кавычек \" на обычные "
foreach ($FirstChangedCode in $tmpMasFirst) {
    $tmpMasSecond += $FirstChangedCode -replace [regex]::Escape($CharBackSlashWhoNeedChange), $CharBackSlashChangable
}

# Экспорт второго этапа обработки
$tmpMasSecond | Out-File -FilePath "$folderPath\$($fileName)_CHGBackSlash.txt" -Force

# Удаление обрамляющих кавычек
foreach ($SecondChangedCode in $tmpMasSecond) {
    $tmpMasThird += $SecondChangedCode.Trim('"')
}

# Финальный экспорт
# --------------------------------------------------
$tmpMasThird | Out-File -FilePath "$folderPath\$($fileName)_CHGBrackets.txt" -Force
Copy-Item -Path "$folderPath\$($fileName)_CHGBrackets.txt" -Destination "$folderPath\$($fileName)_CHGBrackets.csv" -Force

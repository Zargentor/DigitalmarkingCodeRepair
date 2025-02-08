# Скрипт для обработки кодов маркировки из CSV файла
# Основные этапы работы:
# 1. Выбор файла через диалоговое окно
# 2. Предварительная обработка текста
# 3. Экспорт результатов в Excel

# Подключение библиотеки для работы с GUI-элементами
Add-Type -AssemblyName System.Windows.Forms

# Диалоговое окно выбора файла
# --------------------------------------------------
$openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
$openFileDialog.Filter = "Text files (*.csv)|*.csv|All files (*.*)|*.*"  # Фильтр для CSV файлов
$result = $openFileDialog.ShowDialog()  # Показ диалога

# Обработка выбора файла
if ($result -ne [System.Windows.Forms.DialogResult]::OK) {
    Write-Output "Файл не выбран."
    exit
}

# Анализ пути файла
# --------------------------------------------------
$filePath = $openFileDialog.FileName
$folderPath = [System.IO.Path]::GetDirectoryName($filePath)  # Директория файла
$fileName = [System.IO.Path]::GetFileNameWithoutExtension($filePath)  # Имя без расширения
$fileExtension = [System.IO.Path]::GetExtension($filePath)  # Расширение файла

# Чтение и обработка данных
# --------------------------------------------------
$Codes = Get-Content -Path $filePath  # Загрузка исходных данных

# Инициализация временных хранилищ
$tmpMasFirst = @()  # Для кодов до табуляции
$tmpMasSecond = @()  # После замены кавычек
$tmpMasThird = @()  # После удаления обрамляющих кавычек
$tmpMasFourth = @()  # После обрезки по разделителю

# Параметры замены символов
$CharBackSlashWhoNeedChange = '""'  # Исходный паттерн
$CharBackSlashChangable = '"'       # Замена

# Основной цикл обработки
# --------------------------------------------------
foreach ($Code in $Codes) {
    # Извлечение части до табуляции (ASCII 9)
    $tmpMasFirst += $Code.Substring(0, $Code.IndexOf([char]9))
}

# Замена двойных кавычек (экранирование -> обычные)
foreach ($FirstChangedCode in $tmpMasFirst) {
    $tmpMasSecond += $FirstChangedCode -replace [regex]::Escape($CharBackSlashWhoNeedChange), $CharBackSlashChangable
}

# Удаление обрамляющих кавычек
foreach ($SecondChangedCode in $tmpMasSecond) {
    $tmpMasThird += $SecondChangedCode.Trim('"')
}

# Экспорт промежуточных результатов
$tmpMasThird | Out-File -FilePath "$folderPath\$($fileName)_Sliced.csv" -Force

# Обрезка до группового разделителя (ASCII 29)
foreach ($ThirdChangedCode in $tmpMasThird) {
    $tmpMasFourth += $ThirdChangedCode.Substring(0, $ThirdChangedCode.IndexOf([char]29))
}

# Создание Excel-отчёта
# --------------------------------------------------
$excel = New-Object -ComObject Excel.Application
$workbook = $excel.Workbooks.Add()
$worksheet = $workbook.Worksheets.Item(1)

# Запись данных в столбец A
for ($i = 0; $i -lt $tmpMasFourth.Count; $i++) {
    $worksheet.Cells.Item($i+1, 1) = $tmpMasFourth[$i]  # Строки Excel начинаются с 1
}

# Сохранение и завершение работы
$excel.DisplayAlerts = $false  # Игнорировать предупреждения
$workbook.SaveAs("$folderPath\$($fileName)_SlicedShort", 51)  # 51 = xlWorkbookDefault
$workbook.Close($false)
$excel.Quit()

# Освобождение COM-объектов
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[System.GC]::Collect()
[System.GC]::WaitForPendingFinalizers()

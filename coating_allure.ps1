
$testCasesDir = ".\allure-report\data\test-cases"
$outputExcel = ".\scenarios.xlsx"
$sheetName = "Allure"

# === Проверка и подключение ImportExcel ===
if (-not (Get-Module -ListAvailable ImportExcel)) {
    Install-Module -Name ImportExcel -Force -Scope CurrentUser
}
Import-Module ImportExcel -Force


if (-not (Test-Path $testCasesDir)) {
    Write-Error "Папка с тест-кейсами не найдена: $testCasesDir"
    exit 1
}


$rows = @()
$counter = 0

# Рекурсивная функция для поиска блоков "*"
function Get-StarSteps {
    param (
        [array]$steps
    )
    $result = @()
    foreach ($step in $steps) {
        if ($step.name -and $step.name.StartsWith("*")) {
            $result += $step.name.Trim()
        }
        if ($step.steps) {
            $result += Get-StarSteps -steps $step.steps
        }
    }
    return $result
}

# Чтение всех тест-кейсов ===
$jsonFiles = Get-ChildItem -Path $testCasesDir -Filter *.json -Recurse
foreach ($file in $jsonFiles) {
    try {
        $jsonContent = Get-Content $file.FullName -Raw -Encoding UTF8
        $json = $jsonContent | ConvertFrom-Json

        # Сценарии: name не начинается с "*"
        if ($json.name -and -not ($json.name.StartsWith("*"))) {
            $counter++
            $rows += [PSCustomObject]@{
                '№'        = $counter
                'Сценарий' = $json.name
                'Блок'     = ''
            }

            if ($json.testStage.steps) {
                $blocks = Get-StarSteps -steps $json.testStage.steps
                foreach ($block in $blocks) {
                    $rows += [PSCustomObject]@{
                        '№'        = ''
                        'Сценарий' = ''
                        'Блок'     = $block
                    }
                }
            }
        }
    }
    catch {
        Write-Warning "Ошибка при обработке файла $($file.Name): $_"
    }
}

# Экспорт в Excel ===
if ($rows.Count -gt 0) {
    $rows | Export-Excel -Path $outputExcel -WorksheetName $sheetName -AutoSize -BoldTopRow -FreezeTopRow
    Write-Host "Файл успешно создан: $outputExcel" -ForegroundColor Green
} else {
    Write-Host "Нет данных для экспорта" -ForegroundColor Yellow
}
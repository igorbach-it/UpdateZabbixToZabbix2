[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Получение пути к службе Zabbix Agent
$zabbixService = Get-WmiObject win32_service | Where-Object { $_.Name -like '*zabbix*' } | Select-Object -ExpandProperty PathName
$zabbixServiceName = Get-WmiObject win32_service | Where-Object { $_.Name -like '*zabbix*' } | Select-Object -ExpandProperty Name

# Определение пути к каталогу Zabbix Agent
if ($zabbixService -ne $null) {
    $agentPath = $zabbixService -replace '"', '' -replace ' .*', ''
    $agentDirectory = [System.IO.Path]::GetDirectoryName($agentPath)
} else {
    Write-Host "Zabbix Agent не найден."
    exit
}

# Поиск файла конфигурации в каталоге службы
$configFile = Get-ChildItem -Path $agentDirectory -Filter "*.conf" | Where-Object { $_.Name -like "zabbix_*.*" } | Select-Object -ExpandProperty FullName
Write-Host "$configFile"
# Проверка наличия файла конфигурации
if ($configFile -ne $null) {
    $configFilePath = $configFile
} else {
    Write-Host "Файл конфигурации Zabbix Agent не найден."
    exit
}

# Инициализация $newContent текущим содержимым конфигурационного файла
$newContent = Get-Content $configFilePath

# Путь к загрузке новой версии Zabbix Agent 6
$newAgentDownloadURL = "https://cdn.zabbix.com/zabbix/binaries/stable/6.0/6.0.34/zabbix_agent2-6.0.34-windows-amd64-openssl-static.zip"
$tempPath = "C:\Temp"
$newAgentZipPath = Join-Path -Path $tempPath -ChildPath "zabbix_agent2-6.0.34-windows-amd64-openssl-static.zip"
$newAgentExtractPath = Join-Path -Path $tempPath -ChildPath "zabbix_agent2"

# Проверка наличия и создание папки Temp, если она отсутствует
if (-not (Test-Path $tempPath)) {
    New-Item -ItemType Directory -Path $tempPath
}

# Скачивание новой версии Zabbix Agent 6
Invoke-WebRequest -Uri $newAgentDownloadURL -OutFile $newAgentZipPath

# Проверка успешности скачивания
if (-not (Test-Path $newAgentZipPath)) {
    Write-Host "Не удалось скачать файл Zabbix Agent."
    exit
}

# Распаковка архива с новой версией Zabbix Agent 6
Expand-Archive -Path $newAgentZipPath -DestinationPath $newAgentExtractPath

# Проверка наличия распакованных файлов
if (-not (Test-Path "$newAgentExtractPath\bin\zabbix_agent2.exe")) {
    Write-Host "Не удалось распаковать файлы Zabbix Agent."
    exit
}

# Остановка и удаление текущего Zabbix Agent
sc.exe stop "$zabbixServiceName"
Start-Sleep -Seconds 5
sc.exe delete "$zabbixServiceName"

# Копирование новой версии Zabbix Agent
$filesToCopy = @("zabbix_agent2.exe", "zabbix_get.exe", "zabbix_sender.exe")

foreach ($file in $filesToCopy) {
    $newFilePath = Join-Path -Path "$newAgentExtractPath\bin" -ChildPath $file

    if (Test-Path $newFilePath) {
        Write-Host "Копирование $file в $agentDirectory"
        Copy-Item -Path $newFilePath -Destination $agentDirectory -Force
    } else {
        Write-Host "Файл $file не найден в каталоге $newAgentExtractPath."
    }
}

Copy-Item -Path "$newAgentExtractPath\conf\zabbix_agent2.d" -Destination $agentDirectory -Recurse -Force
Copy-Item -Path "$PSScriptRoot\mssql.exe" -Destination $agentDirectory -Force
Copy-Item -Path "$PSScriptRoot\mssql.conf" -Destination $agentDirectory\zabbix_agent2.d\plugins.d\ -Force

# Параметры, которые нужно добавить
$denyKey = "DenyKey=system.run[*]"
$userParameter = "UserParameter=win.description,powershell -NoProfile -ExecutionPolicy Bypass -Command `"`$desc = (Get-CimInstance -ClassName Win32_OperatingSystem).Description; [Console]::OutputEncoding = [System.Text.Encoding]::UTF8; Write-Output `$desc`"" 
$newInclude = "Include=.\zabbix_agent2.d\plugins.d\*.conf"

# Удаление устаревшего параметра EnableRemoteCommands, если он есть
$removeKey = "EnableRemoteCommands=1 "
if ($newContent.Contains($removeKey)) {
    $newContent = $newContent | Where-Object { $_ -ne $removeKey }
    Write-Host "Удалён устаревший параметр EnableRemoteCommands"
}

# Добавление параметра DenyKey, если его нет
if (-not $newContent.Contains($denyKey)) {
    $newContent += "`r`n" + $denyKey
    Write-Host "Добавлен параметр DenyKey"
}

# Добавление параметра UserParameter, если его нет
if (-not $newContent.Contains($userParameter)) {
    $newContent += "`r`n" + $userParameter
    Write-Host "Добавлен параметр UserParameter"
}

# Добавление строки Include для плагинов, если её нет
if (-not $newContent.Contains($newInclude)) {
    $newContent += "`r`n" + $newInclude
    Write-Host "Добавлен параметр Include для плагинов"
}

# Сохранение обновленного файла конфигурации
$newContent | Set-Content $configFilePath


# Запуск Zabbix Agent
& "$agentDirectory\zabbix_agent2.exe" -c "$configFile" --install

Start-Service -Name "Zabbix Agent 2"

# Удаление временных файлов
if (Test-Path "C:\Temp\zabbix*") {
    Remove-Item -Path C:\Temp\zabbix* -Recurse -Force
}

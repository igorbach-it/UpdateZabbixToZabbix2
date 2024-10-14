[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

# Получение пути к службе Zabbix Agent
$zabbixService = Get-WmiObject win32_service | Where-Object { $_.Name -like '*zabbix*' } | Select-Object -ExpandProperty PathName

# Определение пути к каталогу Zabbix Agent
if ($zabbixService -ne $null) {
    # Удаление кавычек и аргументов из строки пути
    $agentPath = $zabbixService -replace '"', '' -replace ' .*', ''
    # Получение только пути к каталогу
    $agentDirectory = [System.IO.Path]::GetDirectoryName($agentPath)
} else {
    Write-Host "$zabbixService не найдена."
    exit
}

# Поиск файла конфигурации в каталоге службы
$configFile = Get-ChildItem -Path $agentDirectory -Filter "*.conf" | Where-Object { $_.Name -like "zabbix_agentd*.conf" } | Select-Object -ExpandProperty FullName

# Проверка наличия файла конфигурации
if ($configFile -ne $null) {
    $configFilePath = $configFile
} else {
    Write-Host "Файл конфигурации Zabbix Agent не найден."
    exit
}

# Вывод используемых путей для проверки
Write-Host "Путь к Zabbix Agent: $agentDirectory"
Write-Host "Путь к файлу конфигурации: $configFilePath"


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

# Остановка Zabbix Agent
#Stop-Service -Name "Zabbix Agent" -ErrorAction SilentlyContinue

sc.exe stop "$zabbixService"
Start-Sleep -Seconds 5


sc.exe delete "$zabbixService"


# Копирование новой версии Zabbix Agent в каталог назначения
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
Copy-Item -Path "$newAgentExtractPath\conf\zabbix_agent2.d\*" -Destination $agentDirectory -Recurse -Force
Copy-Item -Path ".\mssql.exe" -Destination $agentDirectory -Force
Copy-Item -Path ".\mssq.conf" -Destination $newAgentExtractPath\conf\zabbix_agent2.d\plugins.d\ -Force



# Проверка и добавление новых строк в конец файла, если они отсутствуют
$removeKey = "EnableRemoteCommands=1"
$denyKey = "DenyKey=system.run[*]"
$userParameter = "UserParameter=win.description,powershell -NoProfile -ExecutionPolicy Bypass -Command `"`$desc = (Get-CimInstance -ClassName Win32_OperatingSystem).Description; [Console]::OutputEncoding = [System.Text.Encoding]::UTF8; Write-Output `$desc`""
$newInclude = Include=.\zabbix_agent2.d\plugins.d\*.conf

if (-not $content.Contains($denyKey)) {
    $newContent += "`r`n" + $denyKey
}

if (-not $content.Contains($userParameter)) {
    $newContent += "`r`n" + $userParameter
}

if (-not $content.Contains($newInclude)) {
    $newContent += "`r`n" + $newInclude 
}

if ($content.Contains($removeKey)) {  # Закрывающая скобка добавлена
    $newContent = $content | Where-Object { $_ -ne $removeKey }
}

$newContent | Set-Content $configFilePath

# Запуск Zabbix Agent

Start-Process -FilePath "$agentDirectory\zabbix_agent2.exe" -ArgumentList "-c $configFilePath", "--install"


Start-Service -Name "Zabbix Agent 2"

# Удаление временных файлов
Remove-Item -Path C:\Temp\zabbix* -Recurse -Force

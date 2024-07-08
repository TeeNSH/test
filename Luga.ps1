# Включаем поддержку кириллицы в выводе
$OutputEncoding = [Console]::OutputEncoding = [Text.Encoding]::cp866

# Хэш-таблица с именами и адресами устройств
$devices = @{
    "Kassa 6" = "192.168.104.101"
    "Printer 6" = "192.168.104.120"
    "Kassa 10" = "192.168.108.101"
    "Printer 10" = "192.168.108.120"
}

# Пингуем каждое устройство и выводим его доступность
foreach ($deviceName in $devices.Keys) {
    $ipAddress = $devices[$deviceName]
    Write-Host "Пингуем $deviceName с IP-адресом $ipAddress"
    $result = Test-Connection -ComputerName $ipAddress -Count 2 -Quiet
    if ($result) {
        Write-Host "Устройство $deviceName ($ipAddress) доступно" -ForegroundColor Green
    } else {
        Write-Host "Устройство $deviceName ($ipAddress) недоступно" -ForegroundColor Red
    }
}

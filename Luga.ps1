$devices = @{
    "Kassa 6" = "192.168.104.101"
    "Printer 6" = "192.168.104.120"
    "Kassa 10" = "192.168.108.101"
    "Printer 10" = "192.168.108.120"
}

foreach ($deviceName in $devices.Keys) {
    $ipAddress = $devices[$deviceName]
    Write-Host "������� ���������� $deviceName � IP-������� $ipAddress"
    $result = Test-Connection -ComputerName $ipAddress -Count 2 -Quiet
    if ($result) {
        Write-Host "���������� $deviceName ($ipAddress) ��������." -ForegroundColor Green
    } else {
        Write-Host "���������� $deviceName ($ipAddress) ����������." -ForegroundColor Red
    }
}
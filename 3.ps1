# Функция для открытия папки с файлом
function OpenFolder([string]$fullPath) {
    $folderPath = Split-Path -Parent $fullPath
    try {
        Invoke-Item $folderPath
    } catch {
        Write-Warning "Не удалось открыть папку '$folderPath': $_"
    }
}

# Новая функция для копирования файла с новым именем и отображением прогресса
function CopyFileWithNewName($sourceFile) {
    $newFileName = "$(Get-Date -Format 'yyyy-MM-dd') 1Cv8.1CD"
    $destinationFile = Join-Path (Split-Path -Parent $sourceFile) $newFileName
    
    try {
        # Получаем размер файла
        $fileSize = (Get-Item $sourceFile).Length
        $copyBufferSize = 1024 * 512  # Размер буфера копирования (512 КБ)

        # Копируем файл поблочно с отображением прогресса
        $streamReader = New-Object System.IO.FileStream($sourceFile, [System.IO.FileMode]::Open)
        $streamWriter = New-Object System.IO.FileStream($destinationFile, [System.IO.FileMode]::Create)

        $totalBytesRead = 0
        $buffer = New-Object byte[] $copyBufferSize

        while (($bytesRead = $streamReader.Read($buffer, 0, $copyBufferSize)) -gt 0) {
            $streamWriter.Write($buffer, 0, $bytesRead)
            $totalBytesRead += $bytesRead

            # Отображаем прогресс
            $percentComplete = [Math]::Floor(($totalBytesRead / $fileSize) * 100)
            Write-Progress -Activity "Копирование файла" `
                           -Status "Скопировано: $percentComplete%" `
                           -PercentComplete $percentComplete
        }

        $streamReader.Close()
        $streamWriter.Close()

        Write-Host "Файл успешно скопирован: $destinationFile" -ForegroundColor Green
    } catch {
        Write-Warning "Не удалось скопировать файл: $_"
    }
}

# Новая функция для удаления всех файлов и каталогов в указанной папке,
# кроме тех, которые содержат в имени строку "ExtCompT"
function RemoveItemsExceptExtCompT() {
    $targetDir = "$env:USERPROFILE\AppData\Roaming\1C\1cv8"
    $exclusionPattern = "*ExtCompT*"

    try {
        # Удаляем все файлы и каталоги рекурсивно, исключая те, которые соответствуют шаблону исключения
        Get-ChildItem -Path $targetDir -Exclude $exclusionPattern -Recurse | ForEach-Object {
            if ($_.PSIsContainer) {
                Remove-Item -LiteralPath $_.FullName -Recurse -Force
                Write-Host "Удалена папка: $($_.FullName)" -ForegroundColor Yellow
            } else {
                Remove-Item -LiteralPath $_.FullName -Force
                Write-Host "Удален файл: $($_.FullName)" -ForegroundColor Yellow
            }
        }
    } catch {
        Write-Warning "Произошла ошибка при удалении элементов: $_"
    }
}

# Функция для поиска папки "Exchange" на всех локальных дисках
function FindExchangeFolders() {
    # Получаем список всех локальных дисков
    $drives = Get-PSDrive -PSProvider FileSystem |
              Where-Object { $_.DisplayRoot -notlike '\\*' }

    # Создаем пустой массив для хранения результатов
    $results = @()

    foreach ($drive in $drives) {
        # Определяем путь поиска как корень диска
        $searchPath = $drive.Root

        # Ищем папки с именем 'Exchange' на текущем диске
        $folders = Get-ChildItem -Path $searchPath -Filter "Exchange" -Directory -Recurse -ErrorAction SilentlyContinue

        foreach ($folder in $folders) {
            # Добавляем информацию о папке в общий массив
            $results += [PSCustomObject]@{
                Number      = $null
                FolderName  = $folder.Name
                FullPath    = $folder.FullName
            }
        }
    }

    # Присваиваем порядковые номера после сортировки
    for ($i = 0; $i -lt $results.Count; $i++) {
        $results[$i].Number = $i + 1
    }

    # Выводим итоговый результат
    foreach ($result in $results) {
        Write-Host ("{0}. {1}" -f $result.Number, $result.FullPath)
    }

    # Запрашиваем ввод пользователя
    while ($true) {
        $inputNumber = Read-Host "Введите порядковый номер папки для удаления файлов или 'q' для выхода"
        if ($inputNumber -eq 'q') {
            break
        }
        
        if ($inputNumber -match '^\d+$') {
            $index = [int]$inputNumber - 1
            if ($index -ge 0 -and $index -lt $results.Count) {
                $selectedFolder = $results[$index].FullPath
                DeleteInOutFilesFromFolder $selectedFolder
            } else {
                Write-Host "Неверный порядковый номер." -ForegroundColor Yellow
            }
        } else {
            Write-Host "Некорректный ввод." -ForegroundColor Yellow
        }
    }
}

function FindAndDeleteFilesInExchange($path) {
    try {
        # Поиск файлов, содержащих "in" или "out" в имени
        Get-ChildItem -Path $path -Include "*in*", "*out*" -Recurse -Force -File | ForEach-Object {
            Remove-Item -LiteralPath $_.FullName -Force
            Write-Host "Удалённый файл: $($_.FullName)" -ForegroundColor Yellow
        }
    } catch {
        Write-Warning "Произошла ошибка при удалении файлов: $_"
    }
}

# Функция для исправления SyncTrayzor
function Restart-SyncTrayzor {
    # Проверяем, запущен ли процесс syncthing.exe
    $syncthingProcess = Get-Process -Name "syncthing" -ErrorAction SilentlyContinue
    if ($syncthingProcess) {
        # Завершаем процесс
        Stop-Process -InputObject $syncthingProcess -Force
        Write-Host "Процесс syncthing.exe завершен." -ForegroundColor Yellow
    }

    # Проверяем, запущен ли процесс SyncTrayzor.exe
    $syncTrayzorProcess = Get-Process -Name "SyncTrayzor" -ErrorAction SilentlyContinue
    if ($syncTrayzorProcess) {
        # Завершаем процесс
        Stop-Process -InputObject $syncTrayzorProcess -Force
        Write-Host "Процесс SyncTrayzor.exe завершен." -ForegroundColor Yellow
    }

    # Получаем путь к рабочему столу текущего пользователя
    $desktopPath = [Environment]::GetFolderPath("Desktop")

    # Удаляем файл syncthing.exe по пути current user\рабочий стол\SyncTrayzorPortable-x64\data\
    $syncthingDataPath = Join-Path -Path $desktopPath -ChildPath "SyncTrayzorPortable-x64\data"
    $syncthingExePath = Join-Path -Path $syncthingDataPath -ChildPath "syncthing.exe"
    if (Test-Path -Path $syncthingExePath) {
        Rename-Item -Path $syncthingExePath -NewName "syncthing.exe.old"
        Write-Host "Файл syncthing.exe переименован в syncthing.exe.old." -ForegroundColor Yellow
    } else {
        Write-Host "Файл syncthing.exe не найден." -ForegroundColor Red
    }

    # Переименовываем файл syncthing.exe.old обратно в syncthing.exe
    if (Test-Path -Path $syncthingExePath) {
        Rename-Item -Path $syncthingExePath -NewName "syncthing.exe"
        Write-Host "Файл syncthing.exe переименован обратно в syncthing.exe." -ForegroundColor Yellow
    } else {
        Write-Host "Файл syncthing.exe не найден." -ForegroundnRed
    }

    # Формируем полный путь к файлу
    $configFilePath = Join-Path -Path $desktopPath -ChildPath "SyncTrayzorPortable-x64\data\config.xml"

    # Проверяем существование файла
    if (Test-Path -Path $configFilePath) {
        # Удаляем файл
        Remove-Item -Path $configFilePath -Force
        Write-Host "Файл '$configFilePath' успешно удален." -ForegroundColor Green
    } else {
        Write-Host "Файл '$configFilePath' не существует." -ForegroundColor Red
    }

    # Формируем полный путь к программе SyncTrayzor.exe
    $programPath = Join-Path -Path $desktopPath -ChildPath "SyncTrayzorPortable-x64\SyncTrayzor.exe"

    # Проверяем существование программы
    if (Test-Path -Path $programPath) {
        # Запускаем программу
        Start-Process -FilePath $programPath
        Write-Host "Программа '$programPath' запущена." -ForegroundColor Green
    } else {
        Write-Host "Программа '$programPath' не найдена." -ForegroundColor Red
        return
    }

    # Ждем 10 секунд, чтобы дать программе запуститься
    Start-Sleep -Seconds 10

    # Проверяем существование файла
    if (Test-Path -Path $configFilePath) {
        # Читаем содержимое файла
        $xmlContent = [xml](Get-Content -Path $configFilePath)

        # Изменяем значение MinimizeToTray на true
        $xmlContent.SelectSingleNode("//MinimizeToTray").InnerText = "true"

        # Изменяем значение SyncthingDenyUpgrade на true
        $xmlContent.SelectSingleNode("//SyncthingDenyUpgrade").InnerText = "true"

        # Изменяем все значения NotificationsEnabled на false
        $notificationsNodes = $xmlContent.SelectNodes("//NotificationsEnabled")
        foreach ($node in $notificationsNodes) {
            $node.InnerText = "false"
        }

        # Изменяем значение NotifyOfNewVersions на false
        $xmlContent.SelectSingleNode("//NotifyOfNewVersions").InnerText = "false"

        # Изменяем значение DisableHardwareRendering на true
        $xmlContent.SelectSingleNode("//DisableHardwareRendering").InnerText = "true"

        # Изменяем значение EnableFailedTransferAlerts на false
        $xmlContent.SelectSingleNode("//EnableFailedTransferAlerts").InnerText = "false"

        # Изменяем значение EnableConflictFileMonitoring на false
        $xmlContent.SelectSingleNode("//EnableConflictFileMonitoring").InnerText = "false"

        # Изменяем значение PauseDevicesOnMeteredNetworks на false
        $xmlContent.SelectSingleNode("//PauseDevicesOnMeteredNetworks").InnerText = "false"

        # Изменяем значение StartSyncthingAutomatically на true
        $xmlContent.SelectSingleNode("//StartSyncthingAutomatically").InnerText = "true"

        # Изменяем значение MinimizeToTray на true
        $xmlContent.SelectSingleNode("//MinimizeToTray").InnerText = "true"

        # Изменяем значение CloseToTray на true
        $xmlContent.SelectSingleNode("//CloseToTray").InnerText = "true"

        # Изменяем значение ShowDeviceOrFolderRejectedBalloons на false
        $xmlContent.SelectSingleNode("//ShowDeviceOrFolderRejectedBalloons").InnerText = "false"

        # Изменяем значение MinPosition на -32000, -32000
#        $minPositionNode = $xmlContent.SelectSingleNode("//MinPosition")
#        $minPositionNode.InnerText = "-32000,-32000"

        # Сохраняем изменения обратно в файл
        $xmlContent.Save($configFilePath)
        Write-Host "Значения в файле '$configFilePath' изменены." -ForegroundColor Green
    } else {
        Write-Host "Файл '$configFilePath' не существует." -ForegroundColor Red
    }

    # Проверяем, запущен ли процесс syncthing.exe
    $syncthingProcess = Get-Process -Name "syncthing" -ErrorAction SilentlyContinue
    if ($syncthingProcess) {
        # Завершаем процесс
        Stop-Process -InputObject $syncthingProcess -Force
        Write-Host "Процесс syncthing.exe завершен." -ForegroundColor Yellow
    }

    # Проверяем, запущен ли процесс SyncTrayzor.exe
    $syncTrayzorProcess = Get-Process -Name "SyncTrayzor" -ErrorAction SilentlyContinue
    if ($syncTrayzorProcess) {
        # Завершаем процесс
        Stop-Process -InputObject $syncTrayzorProcess -Force
        Write-Host "Процесс SyncTrayzor.exe завершен." -ForegroundColor Yellow
    }
    # Формируем полный путь к программе SyncTrayzor.exe
    $programPath = Join-Path -Path $desktopPath -ChildPath "SyncTrayzorPortable-x64\SyncTrayzor.exe"

    # Проверяем существование программы
    if (Test-Path -Path $programPath) {
        # Запускаем программу
        Start-Process -FilePath $programPath
        Write-Host "Программа '$programPath' запущена." -ForegroundColor Green
    } else {
        Write-Host "Программа '$programPath' не найдена." -ForegroundColor Red
        return
    }
}

# Функция для ЧекДБ файлов 1Cv8.1CD на локальных дисках 
function CheckDB {
    # Меню выбора базы данных
    while ($true) {
        Clear-Host
        Write-Host "Выберите базу данных:" -ForegroundColor Yellow
        Write-Host "1. Федеральная" -ForegroundColor Green
        Write-Host "2. Региональная" -ForegroundColor Green
        Write-Host "Q. Выход" -ForegroundColor Red

        $choice = Read-Host "Введите номер [1-2] или Q для выхода"

        if ($choice -eq 'q') {
            break
        } elseif ($choice -eq '1') {
            $chdbflPath = "C:\Program Files (x86)\1cv8\8.3.14.1993\bin\chdbfl.exe"
            break
        } elseif ($choice -eq '2') {
            $chdbflPath = "C:\Program Files (x86)\1cv8\8.3.10.2561\bin\chdbfl.exe"
            break
        } else {
            Write-Host "Неверный выбор." -ForegroundColor Red
            Start-Sleep -Seconds 2
        }
    }

    if ($choice -ne 'q') {
        # Поиск всех файлов 1Cv8.1CD на всех локальных дисках
        $drives = Get-PSDrive | Where-Object { $_.Provider.Name -eq 'FileSystem' -and $_.Root -like '[A-Z]:\' }
        $files = @()

        foreach ($drive in $drives) {
            $path = $drive.Root + "\"
            Write-Host "Поиск файлов на диске" $path -ForegroundColor Cyan

            try {
                $files += Get-ChildItem -Path $path -Filter "1Cv8.1CD" -Recurse -ErrorAction SilentlyContinue |
                    Select-Object FullName, LastWriteTime
            } catch {}
        }

        if ($files.Count -gt 0) {
            # Вывод списка найденных файлов
            Write-Host "Найденные файлы:" -ForegroundColor Yellow
            for ($i = 0; $i -lt $files.Count; $i++) {
                $ageInDays = ((Get-Date) - $files[$i].LastWriteTime).TotalDays
                if ($ageInDays -lt 7) {
                    Write-Host "$($i+1). $($files[$i].FullName) ($($files[$i].LastWriteTime))" -ForegroundColor Green
                } else {
                    Write-Host "$($i+1). $($files[$i].FullName) ($($files[$i].LastWriteTime))" -ForegroundColor Yellow
                }
            }

            # Запрос ввода номера файла
            while ($true) {
                $fileIndex = Read-Host "Введите номер файла для копирования и запуска программы или Q для выхода"

                if ($fileIndex -eq 'q') {
                    break
                } elseif ([int]$fileIndex -ge 1 -and [int]$fileIndex -le $files.Count) {
                    $selectedFile = $files[[int]$fileIndex - 1]

                    # Создание резервной копии файла
                    $backupFileName = Join-Path (Split-Path $selectedFile.FullName) ("$(Get-Date -Format 'yyyy-MM-dd HH.mm.ss') 1Cv8.1CD")

                    # Отображение статуса копирования
                    $totalSize = (Get-Item $selectedFile.FullName).Length
                    $copiedBytes = 0
                    $progressId = 0

                    try {
                        $streamReader = New-Object System.IO.FileStream $selectedFile.FullName, 'Open', 'Read'
                        $streamWriter = New-Object System.IO.FileStream $backupFileName, 'Create', 'Write'

                        [byte[]]$buffer = New-Object byte[] 4096
                        do {
                            $bytesRead = $streamReader.Read($buffer, 0, $buffer.Length)
                            $streamWriter.Write($buffer, 0, $bytesRead)
                            $copiedBytes += $bytesRead

                            $percentComplete = [Math]::Floor(($copiedBytes / $totalSize) * 100)
                            Write-Progress -Activity "Создание резервной копии" -Status "Копирование файла..." -PercentComplete $percentComplete -Id $progressId
                        } while ($bytesRead -gt 0)

                        $streamReader.Close()
                        $streamWriter.Close()
                    } finally {
                        Remove-Variable streamReader, streamWriter
                    }

                    # Завершение прогресса
                    Write-Progress -Activity "Создание резервной копии" -Completed -Id $progressId

                    # Копирование полного пути до файла в буфер обмена
                    Set-Clipboard -Value $selectedFile.FullName
                    Write-Host "Полный путь к файлу скопирован в буфер обмена: $fullPath" -ForegroundColor Magenta

                    # Получение пути к рабочему столу
                    $desktopPath = [Environment]::GetFolderPath("Desktop")

                    # Имя лог-файла
                    $logFileName = "chdbfl_log_$(Get-Date -Format 'yyyy-MM-dd_HH-mm-ss').txt"
                    $logFilePath = Join-Path $desktopPath $logFileName

                    # Запуск программы chdbfl.exe и запись лога
                    & $chdbflPath $selectedFile.FullName > $logFilePath 2>&1

                    Write-Host "Лог-файл сохранен: $logFilePath" -ForegroundColor Green
                    break
                } else {
                    Write-Host "Номер файла вне диапазона." -ForegroundColor Red
                    Start-Sleep -Seconds 2
                }
            }
        } else {
            Write-Host "Файлы не найдены." -ForegroundColor Red
            Start-Sleep -Seconds 2
        }
    }
}


# Функция для поиска файлов
function SearchFiles() {
    # Получаем список всех локальных дисков
    $drives = Get-PSDrive -PSProvider FileSystem |
              Where-Object { $_.DisplayRoot -notlike '\\*' }

    # Создаем пустой массив для хранения результатов
    $results = @()

    foreach ($drive in $drives) {
        # Определяем путь поиска как корень диска
        $searchPath = $drive.Root

        # Ищем файлы с именем '1Cv8.1CD' на текущем диске
        $files = Get-ChildItem -Path $searchPath -Filter "1Cv8.1CD" -Recurse -ErrorAction SilentlyContinue

        foreach ($file in $files) {
            # Добавляем информацию о файле в общий массив
            $results += [PSCustomObject]@{
                Number      = $null
                FileName    = $file.Name
                FullPath    = $file.FullName
                CreationTime = $file.CreationTime
                LastWriteTime = $file.LastWriteTime
            }
        }
    }

    # Сортируем результаты по дате создания и имени файла
    $sortedResults = $results | Sort-Object -Property CreationTime, FullPath

    # Присваиваем порядковые номера после сортировки
    for ($i = 0; $i -lt $sortedResults.Count; $i++) {
        $sortedResults[$i].Number = $i + 1
    }

    # Определяем возраст файлов относительно текущей даты
    $currentDate = Get-Date
    $oneWeekAgo = $currentDate.AddDays(-7)

    # Выводим итоговый результат
    foreach ($result in $sortedResults) {
        # Определяем цвет вывода в зависимости от возраста файла
        if ($result.LastWriteTime -gt $currentDate) {
            Write-Host "Ошибка: Дата последнего изменения файла в будущем!" -ForegroundColor Red
        } elseif ($result.LastWriteTime -ge $oneWeekAgo) {
            Write-Host ("{0}. {1} ({2})" -f $result.Number, $result.FullPath, $result.LastWriteTime) -ForegroundColor Green
        } else {
            Write-Host ("{0}. {1} ({2})" -f $result.Number, $result.FullPath, $result.LastWriteTime) -ForegroundColor Yellow
        }
    }

    # Запрашиваем ввод пользователя
    while ($true) {
        $inputNumber = Read-Host "Введите порядковый номер файла для выполнения действия или 'q' для выхода"
        if ($inputNumber -eq 'q') {
            break
        }
        
        if ($inputNumber -match '^\d+$') {
            $index = [int]$inputNumber - 1
            if ($index -ge 0 -and $index -lt $sortedResults.Count) {
                $selectedFile = $sortedResults[$index]
                
                # Меняем действие в зависимости от выбранного пункта меню
                switch ($choice) {
                    "1" {
                        OpenFolder $selectedFile.FullPath
                    }
                    "2" {
                        CopyFileWithNewName $selectedFile.FullPath
                    }
                }
            } else {
                Write-Host "Неверный порядковый номер." -ForegroundColor Yellow
            }
        } else {
            Write-Host "Некорректный ввод." -ForegroundColor Yellow
        }
    }
}

# Начальное меню
Write-Host "Добро пожаловать!"
Write-Host "Выберите действие:"
Write-Host "1. Открыть папку с файлом"
Write-Host "2. Скопировать файл с новым именем"
Write-Host "3. Очистить папку %USERPROFILE%\AppData\Roaming\1C\1cv8 (кроме ExtCompT)"
Write-Host "4. Поиск папки 'Exchange' на всех локальных дисках"
Write-Host "5. Исправление SyncTrayzor"
Write-Host "6. CheckDB"
Write-Host "7. ..."
Write-Host "8. ..."
Write-Host "9. ..."
Write-Host "10. Выключение"

do {
    $choice = Read-Host "Введите номер действия"
    switch ($choice) {
        "1" {
            SearchFiles
        }
        "2" {
            SearchFiles
        }
        "3" {
            RemoveItemsExceptExtCompT
        }
        "4" {
            FindExchangeFolders
        }
        "5" {
            Restart-SyncTrayzor
        }
        "6" {
            CheckDB
        }
        "10" {
            Write-Host "До новых встреч!"
            break
        }
        default {
            Write-Host "Недоступно. Выберите действие от 1 до 10."
        }
    }
} until ($choice -eq "10")
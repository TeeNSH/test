# Инициализация пустой строки
$exchangeString = ""

function ShowMenu {
    Clear-Host

    Write-Host "Меню 'Строка обмена'" -ForegroundColor Green
    Write-Host "Текущая строка: $exchangeString" -ForegroundColor Yellow
    Write-Host "1. ФБ" -ForegroundColor Cyan
    Write-Host "2. ООО Продукты 1" -ForegroundColor Cyan
    Write-Host "3. ООО Продукты 2" -ForegroundColor Cyan
    Write-Host "4. Токсово" -ForegroundColor Cyan
    Write-Host "5. Сосново" -ForegroundColor Cyan
    Write-Host "6. Рощино" -ForegroundColor Cyan
    Write-Host "7. Выборг" -ForegroundColor Cyan
    Write-Host "8. Бокситогорск" -ForegroundColor Cyan
    Write-Host "9. Подпорожье" -ForegroundColor Cyan
    Write-Host "10. Кингисепп" -ForegroundColor Cyan
    Write-Host "11. Скопировать итоговую строку" -ForegroundColor Yellow
    Write-Host "12. Очистить итоговую строку" -ForegroundColor Yellow
    Write-Host "13. Выход" -ForegroundColor Red
}

while ($true) {
    ShowMenu

    # Получение выбора пользователя
    $choice = Read-Host "Введите номер пункта"
    
    switch ($choice) {
        1 { $exchangeString += "03-1-##ЦБ-0-0;03-2-##ЦБ-0-0;03-3-##ЦБ-0-0;" }
        2 { $exchangeString += "03-1-##B2-0-0;03-2-##B2-0-0;03-3-##B2-0-0;" }
        3 { $exchangeString += "03-1-##О2-0-0;03-2-##О2-0-0;03-3-##О2-0-0;" }
        4 { $exchangeString += "03-1-##Т1-0-0;03-2-##Т1-0-0;03-3-##Т1-0-0;" }
        5 { $exchangeString += "03-1-##С1-0-0;03-2-##С1-0-0;03-3-##С1-0-0;" }
        6 { $exchangeString += "03-1-##Р1-0-0;03-2-##Р1-0-0;03-3-##Р1-0-0;" }
        7 { $exchangeString += "03-1-##Д1-0-0;03-2-##Д1-0-0;03-3-##Д1-0-0;" }
        8 { $exchangeString += "03-1-##Б1-0-0;03-2-##Б1-0-0;03-3-##Б1-0-0;" }
        9 { $exchangeString += "03-1-##П2-0-0;03-2-##П2-0-0;03-3-##П2-0-0;" }
        10 { $exchangeString += "03-1-##К1-0-0;03-2-##К1-0-0;03-3-##К1-0-0;" }
        11 {
            if ($exchangeString.Length -gt 0) {
                Set-Clipboard -Value $exchangeString
                Write-Host "Итоговая строка скопирована в буфер обмена." -ForegroundColor Green
            } else {
                Write-Host "Нечего копировать, строка пуста." -ForegroundColor Red
            }
        }
        12 {
            $exchangeString = ""
            Write-Host "Итоговая строка очищена." -ForegroundColor Green
        }
        13 { break }
        default { Write-Host "Неправильный выбор." -ForegroundColor Red }
    }

    # Пауза перед следующим циклом
    Start-Sleep -Seconds 1
}
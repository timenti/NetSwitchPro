# Network Adapter Switcher (Portable, Windows)

Портативное приложение для Windows 10/11, которое **реально** переключает состояние двух выбранных сетевых адаптеров:

- отключает активный адаптер;
- включает неактивный адаптер;
- после выполнения обновляет статус в интерфейсе.

Приложение написано на **C# / .NET 8 WinForms** и публикуется в **один `.exe`** (self-contained single-file).

## Что исправлено относительно предыдущей версии

- Полностью удалён legacy-код на TypeScript/Electron, который мог запускаться вместо новой версии и показывать старое сообщение об администраторских правах.
- Логика работы с адаптерами переписана на WMI (`Win32_NetworkAdapter`/`Win32_NetworkAdapterConfiguration`) — без парсинга локализованного вывода `netsh`.
- Включение/отключение выполняется через методы WMI `Enable/Disable` с понятными кодами ошибок.

## Возможности

- Сканирование всех интерфейсов через WMI.
- Экран **System Configuration** в стиле карточек: выбираются ровно 2 адаптера кликом по всей карточке.
- Экран **NetSwitch Controller**: 2 крупные карточки выбранных интерфейсов + центральная круглая кнопка `↻` + кнопка `⚙ Settings`.
- Внизу экрана переключения расположен **лог операций** с автопрокруткой (старт, текущие статусы, выполненные действия, итог).
- Приложение запоминает выбранные 2 адаптера и тему в `settings.json`: при первом запуске открывает экран выбора, при последующих — сразу экран переключения.
- Несколько тем интерфейса: **Dark** и **Light**.
- Отображение имени адаптера, типа, IPv4 и текущего состояния (connected/disabled) с постоянными цветами статусов (зелёный/красный).
- Запрос прав администратора при старте через `app.manifest` (`requireAdministrator`).

## Структура проекта

- `NetworkAdapterSwitcher.csproj` — настройки WinForms и публикации single-file.
- `Program.cs` — точка входа.
- `MainForm.cs` — UI и логика переключения.
- `Services/NetshAdapterService.cs` — WMI-сервис чтения/переключения адаптеров.
- `Models/NetworkAdapterInfo.cs` — модель адаптера.
- `app.manifest` — принудительный запуск с правами администратора.



## Исправления сборки (актуально)

- Устранена ошибка `CS0104` (неоднозначный `Timer`) — в коде используются явные `System.Windows.Forms.Timer`.
- Устранено предупреждение `NETSDK1137` — проект переведён на `Microsoft.NET.Sdk` (при `UseWindowsForms=true`).

## Разрешение конфликтов в вашей ветке

Если GitHub показывает конфликты в `NetworkAdapterSwitcher.csproj`, `README.md`, `Services/NetshAdapterService.cs`, примите версию из этой ветки (она уже без legacy-кода):

```bash
git checkout work
git checkout --theirs NetworkAdapterSwitcher.csproj README.md Services/NetshAdapterService.cs
git add NetworkAdapterSwitcher.csproj README.md Services/NetshAdapterService.cs
git commit -m "Resolve merge conflicts by keeping C# WMI implementation"
```

Если вы делаете merge локально из другой ветки, где остался старый JS/Electron код, безопасная стратегия такая:

1. Оставить C# файлы (`Program.cs`, `MainForm.cs`, `Models/*`, `Services/NetshAdapterService.cs`, `app.manifest`, `.csproj`).
2. Не возвращать удалённые legacy-файлы (`src/*`, `components/*`, `electron/*`, `package.json`, `vite.config.ts`, `tsconfig*.json`).
3. После merge выполнить `dotnet publish` и запускать только `NetworkAdapterSwitcher.exe`.

## Сборка и публикация portable EXE

> Требуется установленный .NET SDK 8.

### 1) Обычная сборка (для проверки кода)

```bash
dotnet build
```

### 2) Публикация одного portable EXE для Windows x64

```bash
dotnet publish -c Release -r win-x64 --self-contained true /p:PublishSingleFile=true
```

Итоговый файл:

```text
bin/Release/net8.0-windows/win-x64/publish/NetworkAdapterSwitcher.exe
```

## Иконка приложения (.ico)

Чтобы после сборки иконка была у `exe` в проводнике и на панели задач:

1. Поместите ваш файл иконки по пути: `Assets/app.ico`.
2. Соберите/опубликуйте проект заново (`dotnet build` или `dotnet publish ...`).

В проекте уже настроено свойство `ApplicationIcon` в `NetworkAdapterSwitcher.csproj`, поэтому иконка будет вшита в исполняемый файл во время сборки.
Дополнительно форма загружает иконку из самого `exe` при старте — это нужно для корректного отображения иконки в заголовке окна и на панели задач.

> Рекомендуется использовать multi-size `.ico` (например 16/24/32/48/256 px), чтобы иконка корректно отображалась в разных местах Windows.


## Где хранится settings.json

Файл настроек больше не создаётся рядом с программой.
Теперь он сохраняется в системную временную папку пользователя:

```text
%TEMP%/NetSwitchPro/settings.json
```

Файл помечается как скрытый (Hidden), чтобы не мешать пользователю.

## Проверка работы

1. Первый запуск: открыть `NetworkAdapterSwitcher.exe`, выбрать 2 адаптера и нажать **Initialize Interface**.
2. Проверить, что на экране переключения отображаются выбранные адаптеры и что в логе появляются записи при нажатии `↻`.
3. Перезапустить приложение: оно должно сразу открыть экран переключения с сохранённой парой адаптеров.
4. При необходимости нажать **⚙ Settings** и выбрать другую пару.
5. Проверить результат также через `ipconfig` и `ncpa.cpl`.

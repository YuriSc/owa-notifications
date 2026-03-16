# OWA Notifier — браузерные уведомления для OWA

Показывает браузерные `Notification API` уведомления о новых письмах пока OWA открыт во вкладке.
Работает через `Office.js makeEwsRequestAsync` — без backend, без CORS проблем.

## Деплой на GitHub Pages

1. Создай новый публичный репозиторий на GitHub (например `owa-notifier`)

2. Загрузи файл `taskpane.html` в корень репозитория

3. Включи GitHub Pages:
   - Settings → Pages → Source: `Deploy from a branch`
   - Branch: `main`, папка: `/ (root)`
   - Сохрани — через ~1 минуту появится URL вида:
     `https://USERNAME.github.io/owa-notifier/`

4. Открой `manifest.xml` и замени в строке `SourceLocation`:
   ```
   https://USERNAME.github.io/REPO/taskpane.html
   ```
   на твой реальный URL, например:
   ```
   https://johndoe.github.io/owa-notifier/taskpane.html
   ```

5. Также замени `<Id>` на уникальный GUID (сгенерируй на https://guidgenerator.com)

## Установка add-in в OWA

1. Открой OWA (`https://owa.mybank.com`)
2. Шестерёнка (Настройки) → Управление надстройками / Manage add-ins
3. Нажми **+** → Добавить из файла
4. Загрузи `manifest.xml`
5. Открой любое письмо — в правой панели появится **OWA Notifier**
6. Нажми **Разрешить уведомления** — браузер запросит разрешение
7. Готово — пока вкладка OWA открыта, уведомления будут приходить каждые 30 секунд

## Как работает

```
taskpane.html (iframe внутри OWA)
  └── каждые 30 сек вызывает Office.context.mailbox.makeEwsRequestAsync()
        └── OWA проксирует запрос к Exchange EWS (без CORS)
              └── Exchange возвращает последние 10 писем из Inbox
                    └── сравниваем с предыдущим состоянием
                          └── новые письма → new Notification(...)
```

## Ограничения

- Уведомления работают **только пока вкладка OWA открыта**
- Только почта (Inbox). Для календаря нужно добавить второй EWS запрос к папке `calendar`
- Интервал проверки — 30 секунд (можно уменьшить до ~10 сек, меньше не рекомендуется)
- Exchange 2016+ / Exchange Online

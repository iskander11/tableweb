# TableWeb — Документация

Корпоративное веб-приложение для совместной работы с Excel-таблицами.

---

## Стек технологий

| Слой | Технология |
|------|-----------|
| Frontend | React 18 + Vite + TypeScript + TailwindCSS |
| Excel-движок | FortuneSheet (open-source) |
| Excel I/O | Python-микросервис (Flask + openpyxl) — высокоточный импорт/экспорт форматирования |
| Backend | Node.js + Express + Socket.io |
| База данных | PostgreSQL |
| Real-time | WebSockets (Socket.io) |
| Сжатие | compression (gzip для HTTP-ответов) |
| Деплой | Nginx + PM2 |

---

## Архитектура

```
tableweb/
├── app/
│   ├── frontend/          # React приложение
│   │   └── src/
│   │       ├── pages/     # LoginPage, DashboardPage, SheetPage, AdminPage
│   │       ├── store/     # Zustand (auth)
│   │       ├── fonts/     # Загрузка пользовательских шрифтов (registry + useFonts)
│   │       └── api/       # axios client
│   ├── backend/           # Node.js API
│   │   └── src/
│   │       ├── routes/    # auth, spreadsheets, excel, backup, fonts
│   │       ├── services/  # excel.js (проксирует в excel-service)
│   │       ├── middleware/ # auth.js (JWT, роли)
│   │       └── db/        # PostgreSQL pool + schema
│   └── excel-service/     # Python (Flask + openpyxl) — парсинг/сборка .xlsx
├── deploy/
│   ├── setup-vps.sh       # Первоначальная установка
│   └── update.sh          # Обновление кода на сервере
└── docs/
    └── README.md
```

---

## Роли пользователей

| Роль | Создание таблиц | Редактирование | Просмотр | Управление пользователями |
|------|:-:|:-:|:-:|:-:|
| **Администратор** | ✅ | ✅ | ✅ | ✅ |
| **Редактор** | ✅ | ✅ | ✅ | ❌ |
| **Читатель** | ❌ | ❌ | ✅ | ❌ |

### Управление пользователями (администратор)

- Логин должен быть уникальным (проверяется без учёта регистра)
- При попытке создать пользователя с уже занятым логином показывается ошибка
- Минимальная длина логина — 2 символа, пароля — 4 символа
- Пользователя можно деактивировать (soft delete), но не удалить полностью
- Администратор может передать свои права другому пользователю (при этом сам становится читателем)

---

## API Endpoints

### Аутентификация
| Метод | URL | Описание |
|-------|-----|----------|
| POST | `/api/auth/login` | Вход в систему |
| GET | `/api/auth/me` | Текущий пользователь |
| POST | `/api/auth/users` | Создать пользователя (admin) |
| GET | `/api/auth/users` | Список пользователей (admin) |
| PATCH | `/api/auth/users/:id/role` | Изменить роль (admin) |
| DELETE | `/api/auth/users/:id` | Деактивировать пользователя (admin) |
| POST | `/api/auth/transfer-admin` | Передать права администратора |

### Таблицы
| Метод | URL | Описание |
|-------|-----|----------|
| GET | `/api/spreadsheets` | Список доступных таблиц |
| POST | `/api/spreadsheets` | Создать таблицу |
| GET | `/api/spreadsheets/:id` | Данные таблицы |
| PATCH | `/api/spreadsheets/:id/rename` | Переименовать |
| DELETE | `/api/spreadsheets/:id` | Удалить |
| POST | `/api/spreadsheets/:id/permissions` | Дать доступ пользователю |
| DELETE | `/api/spreadsheets/:id/permissions/:userId` | Убрать доступ |
| PATCH | `/api/spreadsheets/:id/lock` | Заблокировать/разблокировать |
| PATCH | `/api/spreadsheets/:id/backup-toggle` | Вкл/выкл автобэкап |

### Excel
| Метод | URL | Описание |
|-------|-----|----------|
| POST | `/api/excel/:id/import` | Импорт .xlsx (сохраняет цвета, размеры, формулы) |
| GET | `/api/excel/:id/import-progress` | SSE-поток прогресса импорта |
| GET | `/api/excel/:id/export` | Экспорт в .xlsx |

### Шрифты
| Метод | URL | Описание |
|-------|-----|----------|
| GET | `/api/fonts` | Список шрифтов (для любого авторизованного) |
| POST | `/api/fonts` | Загрузить шрифт .ttf/.otf/.woff/.woff2 (admin) |
| DELETE | `/api/fonts/:id` | Удалить шрифт (admin) |
| GET | `/api/fonts/files/:filename` | Файл шрифта (без авторизации, для `@font-face`) |

### Резервное копирование
| Метод | URL | Описание |
|-------|-----|----------|
| POST | `/api/backup/all` | Создать бэкап всех таблиц (admin) |
| GET | `/api/backup` | История бэкапов (admin) |
| GET | `/api/backup/:id/download` | Скачать бэкап (admin) |
| DELETE | `/api/backup/:id` | Удалить бэкап (admin) |
| POST | `/api/backup/:id/restore` | Восстановить таблицы из бэкапа (admin) |

---

## WebSocket события

| Событие | Направление | Описание |
|---------|-------------|----------|
| `join-sheet` | client → server | Войти в комнату таблицы |
| `cell-change` | оба | Изменение ячейки (real-time синхронизация) |
| `save-sheet` | client → server | Сохранить состояние в БД |
| `room-users` | server → client | Список кто сейчас в таблице |

---

## Резервное копирование

### Ручной бэкап
Администратор может создать бэкап всех таблиц в разделе **Настройки → Резервные копии**. Бэкап сохраняется как ZIP-архив с .xlsx файлами.

### Автоматический бэкап
Каждое воскресенье в 02:00 сервер автоматически создаёт ZIP-архив таблиц, у которых включён автобэкап. Архивы хранятся в директории `BACKUP_DIR` (по умолчанию `./backups/`).

### Включение автобэкапа для таблицы
На главной странице (список таблиц) у каждой таблицы есть иконка архива (📦). Нажмите на неё:
- **Зелёная** — автобэкап включён
- **Серая** — автобэкап выключен

### Восстановление из бэкапа
В разделе **Настройки → Резервные копии** можно восстановить таблицы из любого бэкапа. Таблицы создаются заново (с суффиксом «восстановлено» при совпадении имён).

---

## Первоначальная установка на VPS

```bash
# 1. Скопировать setup-vps.sh на сервер
scp deploy/setup-vps.sh root@<IP>:/root/

# 2. Подключиться к серверу
ssh root@<IP>

# 3. Открыть скрипт и заменить:
#    - CHANGE_THIS_PASSWORD  → пароль для PostgreSQL
#    - CHANGE_THIS_TO_RANDOM_SECRET → случайная строка для JWT
#    - TOKEN_HERE → GitHub Personal Access Token

# 4. Запустить установку
chmod +x /root/setup-vps.sh
bash /root/setup-vps.sh
```

После установки: **http://<IP>**
Первый вход: `admin` / `admin123` — **сменить пароль сразу!**

---

## Обновление кода

```bash
ssh root@<IP>
bash /var/www/tableweb/deploy/update.sh
```

---

## Передача прав администратора

1. Войти как администратор
2. Перейти в **Настройки** (кнопка в шапке)
3. Раздел «Передать права администратора»
4. Выбрать пользователя → нажать **Передать**
5. Вы автоматически становитесь читателем и выходите из системы

---

## Импорт/Экспорт Excel

**Что сохраняется при импорте:**
- Данные ячеек
- Формулы (с результатами)
- Цвета ячеек (заливка и цвет текста), включая условное форматирование
- Размеры колонок и строк
- Объединённые ячейки (merge) — мастер + слейв-ячейки `mc`
- Шрифт: название (как в Excel), жирный, курсив, размер, подчёркивание
- Выравнивание и перенос текста (`tb`)
- Поворот текста: вертикальный и под углом (`tr`/`rt`)
- Автофильтр — стрелки на строке заголовка (видны редакторам)
- Границы ячеек
- Несколько листов

**При экспорте** — все те же данные записываются обратно в .xlsx формат.

---

## Пользовательские шрифты

Чтобы импортированная таблица выглядела 1:1 с оригиналом, нужный шрифт (например,
Times New Roman) должен быть доступен браузеру. Администратор может загружать шрифты:

1. **Настройки → Шрифты** → выбрать файл `.ttf/.otf/.woff/.woff2`.
2. Имя шрифта определяется из файла автоматически (можно задать вручную, если нужно точное совпадение с Excel).
3. После загрузки шрифт:
   - применяется при отрисовке импортированных таблиц (через `FontFace` / `@font-face`);
   - появляется в выпадающем списке шрифтов в панели инструментов таблицы.

**Технически:** файлы лежат в `FONTS_DIR` (по умолчанию `./uploads/fonts`) и отдаются
статикой по `/api/fonts/files/...` без авторизации (иначе браузер не сможет подгрузить
шрифт). Метаданные — в таблице `fonts`.

---

## Оптимизация производительности

- **gzip-сжатие** — все HTTP-ответы сжимаются middleware `compression`, что значительно ускоряет загрузку больших таблиц (JSONB данные хорошо сжимаются)
- **Debounce сохранения** — изменения сохраняются в БД не чаще чем раз в 2 секунды, что снижает нагрузку при быстром редактировании
- **Индексы БД** — добавлены индексы на часто используемые поля (spreadsheet_data, permissions, users)
- **Кеширование шрифтов** — файлы шрифтов отдаются с заголовком `Cache-Control: max-age=7d, immutable`

---

## База данных

### Применение индексов (при обновлении)

Индексы описаны в `schema.sql`. Для применения на существующей базе:

```sql
CREATE INDEX IF NOT EXISTS idx_spreadsheet_data_sheet_id ON spreadsheet_data(spreadsheet_id);
CREATE INDEX IF NOT EXISTS idx_spreadsheet_permissions_sheet_id ON spreadsheet_permissions(spreadsheet_id);
CREATE INDEX IF NOT EXISTS idx_spreadsheet_permissions_user_id ON spreadsheet_permissions(user_id);
CREATE INDEX IF NOT EXISTS idx_users_username_lower ON users(LOWER(username));
CREATE INDEX IF NOT EXISTS idx_spreadsheets_created_by ON spreadsheets(created_by);
CREATE INDEX IF NOT EXISTS idx_backups_created_at ON backups(created_at DESC);
```

# TableWeb — Документация

Корпоративное веб-приложение для совместной работы с Excel-таблицами.

---

## Стек технологий

| Слой | Технология |
|------|-----------|
| Frontend | React 18 + Vite + TypeScript + TailwindCSS |
| Excel-движок | FortuneSheet (open-source) |
| Excel I/O | ExcelJS (полный импорт/экспорт с форматированием) |
| Backend | Node.js + Express + Socket.io |
| База данных | PostgreSQL |
| Real-time | WebSockets (Socket.io) |
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
│   │       └── api/       # axios client
│   └── backend/           # Node.js API
│       └── src/
│           ├── routes/    # auth, spreadsheets, excel, backup
│           ├── services/  # excel.js (импорт/экспорт)
│           └── db/        # PostgreSQL pool + schema
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
| GET | `/api/excel/:id/export` | Экспорт в .xlsx |

### Резервное копирование
| Метод | URL | Описание |
|-------|-----|----------|
| POST | `/api/backup/all` | Создать бэкап всех таблиц (admin) |
| GET | `/api/backup` | История бэкапов (admin) |

---

## WebSocket события

| Событие | Направление | Описание |
|---------|-------------|----------|
| `join-sheet` | client → server | Войти в комнату таблицы |
| `cell-change` | оба | Изменение ячейки (real-time синхронизация) |
| `save-sheet` | client → server | Сохранить состояние в БД |
| `room-users` | server → client | Список кто сейчас в таблице |

---

## Первоначальная установка на VPS

```bash
# 1. Скопировать setup-vps.sh на сервер
scp deploy/setup-vps.sh root@168.222.202.6:/root/

# 2. Подключиться к серверу
ssh root@168.222.202.6

# 3. Открыть скрипт и заменить:
#    - CHANGE_THIS_PASSWORD  → пароль для PostgreSQL
#    - CHANGE_THIS_TO_RANDOM_SECRET → случайная строка для JWT
#    - TOKEN_HERE → GitHub Personal Access Token

# 4. Запустить установку
chmod +x /root/setup-vps.sh
bash /root/setup-vps.sh
```

После установки открыть: **http://168.222.202.6**  
Первый вход: `admin` / `admin123` — **сменить пароль сразу!**

---

## Обновление кода

```bash
ssh root@168.222.202.6
bash /var/www/tableweb/deploy/update.sh
```

---

## Передача прав администратора

1. Войти как администратор
2. Перейти в **Настройки** (кнопка в шапке)
3. Раздел "Передать права администратора"
4. Выбрать пользователя → нажать **Передать**
5. Вы автоматически становитесь читателем и выходите из системы

---

## Импорт/Экспорт Excel

**Что сохраняется при импорте:**
- ✅ Данные ячеек
- ✅ Формулы (с результатами)
- ✅ Цвета ячеек (заливка и цвет текста)
- ✅ Размеры колонок и строк
- ✅ Объединённые ячейки (merge)
- ✅ Шрифт (жирный, курсив, размер)
- ✅ Выравнивание и перенос текста
- ✅ Границы ячеек
- ✅ Несколько листов

**При экспорте** — все те же данные записываются обратно в .xlsx формат.

---

## Автоматическое резервное копирование

Каждое воскресенье в 02:00 сервер автоматически создаёт ZIP-архив со всеми таблицами, у которых включён бэкап. Архивы хранятся в `/var/www/tableweb/backups/`.

Включить бэкап для таблицы: кнопка в списке таблиц (только для владельца или admin).

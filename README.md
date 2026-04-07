# TOM docx-service

Микросервис генерации должностных инструкций для TOM LLC.  
Принимает структурированный JSON → возвращает готовый `.docx` файл.

## Деплой на Railway (5 минут)

1. Создай репозиторий на GitHub, загрузи все файлы этой папки
2. Зайди на [railway.app](https://railway.app) → **New Project** → **Deploy from GitHub repo**
3. Выбери репозиторий → Railway автоматически определит Node.js и задеплоит
4. Зайди в **Settings → Networking → Generate Domain** — получишь URL вида `https://tom-docx-service-production.up.railway.app`

## API endpoints

### GET /health
Проверка работоспособности.
```json
{ "status": "ok", "service": "TOM docx-service", "version": "1.0.0" }
```

### POST /parse
Извлечение текста из `.docx` анкеты (для n8n → Claude parse step).

**Body:**
```json
{
  "base64": "<base64-encoded docx file content>"
}
```

**Response:**
```json
{
  "text": "Полный текст анкеты...",
  "messages": []
}
```

### POST /generate
Генерация должностной инструкции.

**Body** (полная структура JD данных):
```json
{
  "pos": {
    "en": "Chemical Engineer",
    "ru": "Инженер-химик",
    "uz": "Muxandis kimyogar",
    "dept_en": "Chemical Water Treatment Workshop",
    "dept_ru": "Цех химической водоподготовки",
    "dept_uz": "Suvni kimyoviy tozalash sexi"
  },
  "docCode": "TOM_INT_HRD_JD_SKTS_006",
  "filename": "TOM_INT_HRD_JD_SKTS_006_EMP_Chemical_Engineer_EN_RU_UZ_V1.docx",
  "reportsTo": {
    "en": "Laboratory Supervisor",
    "ru": "Начальник лаборатории",
    "uz": "Laboratoriya boshlig'i"
  },
  "directReports": "None / Нет / Yo'q",
  "grade": "GSO-4 | Specialist",
  "purpose": {
    "en": "The Chemical Engineer performs...",
    "ru": "Инженер-химик выполняет...",
    "uz": "Muxandis kimyogar amalga oshiradi..."
  },
  "duties": [
    {
      "title": {
        "en": "3.1 Water Quality Control (20%)",
        "ru": "3.1 Контроль качества воды (20%)",
        "uz": "3.1 Suv sifati nazorati (20%)"
      },
      "en": ["Monitor compliance with schedules", "Report deviations"],
      "ru": ["Контролировать соблюдение графиков", "Докладывать об отклонениях"],
      "uz": ["Jadvallar bajarilishini nazorat qilish", "Og'ishlarni bildirish"]
    }
  ],
  "kpis": [
    ["Analysis schedule compliance", "Выполнение графика", "Jadval bajarilishi", "100%", "Monthly"]
  ],
  "education": {
    "en": "Higher non-specialised education in chemistry.",
    "ru": "Высшее (не профильное) образование в области химии.",
    "uz": "Kimyo sohasida yuqori ta'lim."
  },
  "expYears": 3,
  "languages": {
    "en": "Intermediate Russian. MS Office.",
    "ru": "Средний уровень русского языка. MS Office.",
    "uz": "O'rta darajadagi rus tili. MS Office."
  },
  "schedule": {
    "en": "5-day working week, Monday to Friday.",
    "ru": "Пятидневная рабочая неделя.",
    "uz": "Dushanba-juma, 5 kunlik ish haftasi."
  },
  "career": {
    "upward": {
      "en": "Laboratory Supervisor",
      "enDesc": "With 3+ years of demonstrated excellence.",
      "ru": "Начальник лаборатории",
      "ruDesc": "При 3+ годах высоких результатов.",
      "uz": "Laboratoriya boshlig'i",
      "uzDesc": "3+ yil davomida yuqori samaradorlik bilan."
    },
    "lateral": null
  },
  "supervisor": {
    "en": "Laboratory Supervisor",
    "ru": "Начальник лаборатории",
    "uz": "Laboratoriya boshlig'i"
  }
}
```

**Response:** бинарный `.docx` файл с заголовком `Content-Disposition: attachment`.

## Локальный запуск

```bash
npm install
npm run dev
# Сервис доступен на http://localhost:3000
```

## Тест через curl

```bash
# Health check
curl https://YOUR_RAILWAY_URL/health

# Parse анкеты
curl -X POST https://YOUR_RAILWAY_URL/parse \
  -H "Content-Type: application/json" \
  -d '{"base64": "'$(base64 -i анкета.docx)'"}'

# Генерация ДИ
curl -X POST https://YOUR_RAILWAY_URL/generate \
  -H "Content-Type: application/json" \
  -d @payload.json \
  --output result.docx
```

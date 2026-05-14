# بوابة جامعة طرابلس — Tripoli University Portal

## Quick Start

### 1. Start the database
```bash
docker compose up -d
```

### 2. Install Python dependencies
```bash
pip install -r requirements.txt
```

### 3. Seed the database
```bash
python seed.py
```

### 4. Run the app
```bash
python app.py
```

Open http://localhost:5000



## Features

```
student_managment_system/
├── app.py                  # Flask routes and business logic
├── models.py               # SQLAlchemy database models
├── config.py               # App configuration
├── seed.py                 # Database seeding script
├── requirements.txt
├── docker-compose.yml      # PostgreSQL container
└── templates/
    ├── base.html
    ├── auth/
    │   ├── login.html
    │   └── signup.html
    ├── student/
    │   ├── dashboard.html
    │   ├── enrollment.html
    │   └── enrollment_locked.html
    └── admin/
        ├── dashboard.html
        ├── students.html
        └── pre_register.html
```

## Environment Variables

| Variable | Default | Description |
|----------|---------|-------------|
| `SECRET_KEY` | (dev key) | Flask secret key — change in production |
| `DATABASE_URL` | postgresql://portal_user:portal_pass@localhost:5432/tripoli_portal | DB connection string |

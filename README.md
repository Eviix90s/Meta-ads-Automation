# Meta Ads Automation Platform

> Automated web platform for managing Meta (Facebook/Instagram) advertising campaigns — built for internal business operations at **Vuela a la Vida**.

---

## Overview

This project started as a **standalone Python script** that automated repetitive Meta Ads tasks through the browser. It had no UI, no authentication, and no persistence — just raw automation running locally on one machine.

Over several development iterations, it evolved into a **full-stack web platform** with:

- A modern, branded web interface
- Secure user authentication backed by PostgreSQL
- Containerized deployment with Docker
- Automated CI/CD pipeline via GitHub Actions

---

## Evolution

| Version | What it was |
|---------|-------------|
| v1.0 | Python automation script — no UI, runs locally |
| v2.0 | NiceGUI web interface — accessible from browser, no auth |
| v3.0 | Login system with animated splash screen + session management |
| v3.1 | PostgreSQL authentication, Docker containerization, CI/CD pipeline |

---

## Features

- **Animated Login Page** — glassmorphism card, video background, splash → form transition
- **PostgreSQL Authentication** — bcrypt-hashed passwords, email domain validation (`@vuelaalavida.com`)
- **Session Management** — NiceGUI user storage, protected routes, logout button
- **Meta Ads Automation** — automated campaign management via browser automation
- **Google Sheets Integration** — reads and writes campaign data to spreadsheets
- **Dark / Light Mode** — theme toggle with persistent preference
- **Docker + CI/CD** — push to `main` → server auto-deploys via GitHub Actions self-hosted runner

---

## Tech Stack

| Layer | Technology |
|-------|-----------|
| Frontend | NiceGUI (Python-based reactive UI) |
| Backend | FastAPI + Python 3.13 |
| Database | PostgreSQL 17 |
| Auth | bcrypt password hashing |
| Automation | Browser automation + Meta Ads API |
| Sheets | Google Sheets API (gspread) |
| Containerization | Docker + Docker Compose |
| CI/CD | GitHub Actions (self-hosted runner on Windows server) |
| Tunnel | Cloudflare Tunnel (HTTPS, no open ports) |

---

## Architecture

```
Developer PC
     │
     │  git push → GitHub
     │
     ▼
GitHub Actions
     │
     │  triggers self-hosted runner on server
     │
     ▼
Windows Server (24/7)
  ├── GitHub Actions Runner  ← listens for deploy jobs
  ├── Docker Container       ← app_web.py (NiceGUI + FastAPI)
  ├── PostgreSQL 17 (native) ← user authentication
  └── Cloudflare Tunnel      ← public HTTPS access, no open ports
```

---

## Security

- No credentials committed to the repository
- All secrets managed via `.env` file (excluded from git)
- Database password, DB user, and session secret loaded from environment variables
- `credenciales.json` (Google service account) excluded from git, mounted as Docker volume
- Access restricted to users with `@vuelaalavida.com` email domain

---

## Setup

### Prerequisites
- Docker Desktop
- PostgreSQL 17
- Python 3.13 (for local development)

### Local Development
```bash
# Clone the repo
git clone https://github.com/Eviix90s/Meta-ads-Automation.git
cd Meta-ads-Automation

# Create environment file
cp .env.example .env
# Edit .env with your database credentials

# Install dependencies
pip install -r requirements.txt

# Run
python app_web.py
```

### Production (Docker)
```bash
# Build and start
docker compose up -d --build

# View logs
docker compose logs -f
```

---

## CI/CD Flow

1. Push code to `main` branch
2. GitHub Actions triggers the `Deploy to Server` workflow
3. Self-hosted runner on the Windows server picks up the job
4. Runner runs `docker compose down && docker compose up -d --build`
5. New version is live — zero manual intervention

---

## Screenshots

> Login page with animated video background and glassmorphism card

> Main dashboard with Meta Ads automation controls

> Dark mode support

---

*Built with Python · NiceGUI · PostgreSQL · Docker · GitHub Actions*

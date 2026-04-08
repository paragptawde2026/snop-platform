# SNOP Platform v2 — Enterprise Olefin Operations Planning

Production-grade Sales & Operations Planning system for multi-plant olefin crackers.
**FastAPI · PostgreSQL · Celery · React · Docker · Pyomo MINLP**

---

## What This System Does (Plain English)

| Feature | What it means |
|---------|---------------|
| **Multi-plant** | Manage 4+ crackers, each with its own furnaces and feed data |
| **Furnace management** | Track each furnace — coil type, capacity, status, decoking schedule |
| **Yield matrix** | Upload historical Excel data; the system builds a per-feed yield table |
| **MINLP optimization** | Solves which feedstock each furnace should crack to maximise profit |
| **Scenario planning** | Create what-if cases, change prices or constraints, re-run the solver |
| **Excel upload** | Drop any .xlsx or .csv — column names are auto-detected |

---

## Architecture

```
Browser (React) ──► Nginx ──► FastAPI (8000)
                              │
                              ├── PostgreSQL  (data store)
                              ├── Redis       (Celery broker)
                              └── Celery Worker (MINLP solver)
```

### MINLP Solver Chain

The optimizer tries solvers in this order:
1. **GLPK** (installed in Docker image — free, exact MIP)
2. **BONMIN** (COIN-OR MINLP — if available)
3. **COUENNE** (global MINLP — if available)
4. **IPOPT** (NLP relaxation — if available)
5. **scipy HiGHS** (LP relaxation — always available, fast fallback)

The binary variable `y[f,k] ∈ {0,1}` forces each furnace to crack exactly one feedstock — this is the key MINLP constraint.

---

## Quick Start

### Step 1 — Prerequisites

Install these two things (both free):
- **Docker Desktop**: https://www.docker.com/products/docker-desktop
- **Git**: https://git-scm.com/downloads

### Step 2 — Setup (run once)

Open a terminal (Command Prompt on Windows, Terminal on Mac/Linux):

```bash
# Go to your project folder
cd C:\Users\yourname\Desktop\snop       # Windows
cd ~/Desktop/snop                        # Mac/Linux

# Create your settings file
copy .env.example .env                  # Windows
cp .env.example .env                    # Mac/Linux
```

Open the `.env` file in Notepad and change these three lines:
```
POSTGRES_PASSWORD=my_strong_password_here
SECRET_KEY=any_random_32_characters_here
```

### Step 3 — Start

```bash
docker-compose up -d
```

Wait about 60 seconds for everything to start, then open:
- **Platform**: http://localhost:3000
- **API docs**: http://localhost:8000/api/docs

### Step 4 — First-time configuration

1. **Add plants** → Plants → Add Plant
2. **Add furnaces** → Furnaces → Add Furnace (assign to plant)
3. **Add feedstocks** via API: `POST http://localhost:8000/api/feeds`
4. **Upload yield data** → Data Upload → Yield Data
5. **Upload product prices** → Data Upload → Product Prices
6. **Create scenario** → Scenarios → New Scenario → Optimise

---

## Excel Upload Format

All files accept `.xlsx`, `.xls`, `.csv`. Column names are auto-detected.

### Yield Data
| FurnaceCode | FeedCode | Date | Ethylene% | Propylene% | Butadiene% | Benzene% | COT |
|---|---|---|---|---|---|---|---|
| F-A1 | ETH | 2024-01-15 | 52.3 | 3.1 | 1.5 | 0.8 | 845 |
| F-A1 | LPG | 2024-01-15 | 38.2 | 14.5 | 4.2 | 2.1 | 820 |

### Feed Prices
| FeedCode | Date | Price_USD_T |
|---|---|---|
| ETH | 2024-07-01 | 310 |

### Product Prices
| Product | Date | Price_USD_T |
|---|---|---|
| ethylene | 2024-07-01 | 1250 |
| propylene | 2024-07-01 | 980 |

---

## Daily Commands

```bash
# Start the platform
docker-compose up -d

# Stop the platform (keeps your data)
docker-compose down

# Check all services are running
docker-compose ps

# View backend logs
docker-compose logs -f backend

# View worker (optimizer) logs
docker-compose logs -f worker

# Restart just the backend
docker-compose restart backend
```

---

## Troubleshooting

| Problem | Fix |
|---------|-----|
| Page won't load | Check Docker Desktop is running. Run `docker-compose ps` and look for `Up`. |
| "Unknown furnace" in upload | Furnace code in Excel must exactly match what's in the database |
| Optimization says "infeasible" | Check: yield data uploaded for the date range, product prices exist, furnaces set to "running" |
| Optimization uses LP not MINLP | GLPK is installed in the container. Check worker logs: `docker-compose logs worker` |
| Frontend shows blank | Wait 30s after starting. Run `docker-compose restart frontend` |

---

## Backup & Restore

```bash
# Backup your database
docker-compose exec db pg_dump -U snop_admin snop_production > backup_$(date +%Y%m%d).sql

# Restore
docker-compose exec -T db psql -U snop_admin snop_production < backup_20240701.sql
```

---

## Production Deployment

For a real server, also set in `.env`:
```
ENVIRONMENT=production
CORS_ORIGINS=https://yourdomain.com
VITE_API_BASE_URL=https://api.yourdomain.com
```

Generate a strong SECRET_KEY:
```bash
python3 -c "import secrets; print(secrets.token_hex(32))"
```

---

## File Structure

```
snop/
├── docker-compose.yml          # All 6 services
├── .env.example                # Settings template
├── docker/
│   ├── nginx/nginx.conf        # Reverse proxy
│   └── postgres/init.sql       # DB extensions
├── backend/
│   ├── Dockerfile
│   ├── requirements.txt        # Python deps incl. Pyomo
│   └── app/
│       ├── main.py             # FastAPI entry point
│       ├── core/               # Config, DB, Celery
│       ├── models/models.py    # All 11 DB tables
│       ├── schemas/schemas.py  # Pydantic validation
│       ├── services/
│       │   ├── optimizer.py    # MINLP + LP solver
│       │   ├── excel_service.py# Upload parser
│       │   └── optimization_tasks.py  # Celery task
│       └── api/routes/         # 9 route files
└── frontend/
    ├── Dockerfile
    ├── src/
    │   ├── App.tsx             # Router
    │   ├── api/client.ts       # Typed API layer
    │   ├── components/         # Sidebar, UI primitives
    │   └── pages/              # 7 full pages
    └── tailwind.config.js
```

---

## Tech Stack Versions

| Component | Version |
|-----------|---------|
| Python | 3.12 |
| FastAPI | 0.111 |
| SQLAlchemy | 2.0 (async) |
| Pyomo | 6.7 |
| GLPK | system package |
| PostgreSQL | 16 |
| Redis | 7 |
| Celery | 5.3 |
| React | 18 |
| Vite | 5 |
| TailwindCSS | 3.4 |
| Node.js | 20 |

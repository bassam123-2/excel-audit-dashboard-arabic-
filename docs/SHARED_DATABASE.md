# Shared MySQL database (team)

All developers can use **one MySQL server** so upload metadata (`UploadSession`, `ExportLog`, `EmailLog`) is shared.

Excel files themselves still live on each machine under `uploads/` unless you add shared file storage later.

## Option A — Docker on one team machine (recommended for dev)

Whoever hosts the DB (your PC or a small server):

1. Install [Docker Desktop](https://www.docker.com/products/docker-desktop/).
2. Copy `.env.example` to `.env` and set strong passwords:

```env
MYSQL_ROOT_PASSWORD=choose-a-strong-root-password
MYSQL_PASSWORD=choose-a-strong-app-password
MYSQL_USER=excel_user
MYSQL_DATABASE=excel_arabic
MYSQL_PORT=3306
```

3. Start the database:

```bash
docker compose up -d
```

4. Share with the team (secure channel, **not** GitHub):
   - Host IP or hostname (e.g. `192.168.1.50` or `db.mycompany.local`)
   - `MYSQL_USER`, `MYSQL_PASSWORD`, `MYSQL_DATABASE`, port `3306`

Each developer puts the same values in their local `.env`:

```env
MYSQL_HOST=192.168.1.50
MYSQL_PORT=3306
MYSQL_DATABASE=excel_arabic
MYSQL_USER=excel_user
MYSQL_PASSWORD=<shared password from team lead>
```

5. On each machine, after `pip install -r requirements.txt`:

```bash
python manage.py migrate
python manage.py runserver 0.0.0.0:8765
```

Only **one person** should run migrations when models change; everyone else runs `migrate` after pulling.

## Option B — Company MySQL server (production / staging)

Use your IT-provided host (Azure MySQL, AWS RDS, on-prem MySQL, etc.):

```env
MYSQL_HOST=your-mysql-server.company.com
MYSQL_PORT=3306
MYSQL_DATABASE=excel_arabic
MYSQL_USER=excel_arabic_app
MYSQL_PASSWORD=<from IT>
# If SSL is required:
MYSQL_REQUIRE_SSL=1
# MYSQL_SSL_CA=C:\path\to\ca.pem
```

Ask IT to create the database with `utf8mb4` / `utf8mb4_unicode_ci`.

## Option C — Local MySQL only (solo)

```env
MYSQL_HOST=127.0.0.1
MYSQL_USER=root
MYSQL_PASSWORD=your-local-password
```

## Firewall

If others cannot connect:

- Windows: allow inbound **TCP 3306** on the host running Docker MySQL.
- Docker: port `3306:3306` is already published in `docker-compose.yml`.

Test from another PC:

```bash
mysql -h 192.168.1.50 -P 3306 -u excel_user -p excel_arabic
```

## What is stored in the shared DB

| Table | Purpose |
|-------|---------|
| `uploads_uploadsession` | Who uploaded which file, path, hash, time |
| `uploads_exportlog` | PPTX/HTML export attempts |
| `uploads_emaillog` | Email send attempts |

## Security

- Never commit `.env` with real passwords.
- Use different passwords for **dev shared** vs **production**.
- Restrict MySQL user to one database and only needed privileges.

## Troubleshooting

| Problem | Fix |
|---------|-----|
| `Can't connect to MySQL server` | Check `MYSQL_HOST`, firewall, Docker running |
| `Access denied` | Wrong user/password; recreate user in MySQL |
| `Unknown database` | Run `docker compose up -d` or create DB manually |
| Migrations out of sync | Pull latest code, run `python manage.py migrate` |

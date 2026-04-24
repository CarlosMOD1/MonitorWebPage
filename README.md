# Failure Monitoring Dashboard — Deployment Guide
## Jesus esta colaborando
## Prerequisites

- Python 3.10+: https://www.python.org/downloads/
- [ODBC Driver 17 or 18 for SQL Server](https://learn.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server)

---

## 1. Install Dependencies

```bash
pip install flask pyodbc pandas waitress
```

---

## 2. Configure Database Connection

Create a `.env` file in the project root with your SQL Server credentials:

```
DB_SERVER=YOUR_SERVER_IP
DB_NAME=YOUR_DB
DB_USER=YOUR_USER
DB_PASS=YOUR_PASSWORD
```

`app.py` reads these values automatically via `python-dotenv`. Never commit the `.env` file to version control.

---

## 3. Configure for Production

In `app.py`, make sure the last line uses `host='0.0.0.0'` so the app is accessible from the network:

```python
app.run(host='0.0.0.0', port=5000)
```

### Recommended: Use Waitress (production WSGI server)

Create a file `serve.py` in the project folder:

```python
from waitress import serve
from app import app

serve(app, host='0.0.0.0', port=5000, threads=4)
```

Run `serve.py` instead of `app.py` for production.

---

## 4. Open Firewall Port

Run the following command as Administrator to allow inbound traffic on port 5000:

```bash
netsh advfirewall firewall add rule name="Failure Monitoring" dir=in action=allow protocol=TCP localport=5000
```

---

## 5. Auto-start on Boot (Task Scheduler)

1. Open **Task Scheduler** → Create Basic Task
2. **Trigger:** When the computer starts
3. **Action:** Start a program
   - Program: `python`
   - Arguments: `C:\Users\JesusMoyaLozano\Desktop\Cowork\MonitorWebPage\serve.py`
4. Check **"Run whether user is logged on or not"**
5. Check **"Run with highest privileges"**

---

## 6. Access the Dashboard

From any computer on the same network, open a browser and go to:

```
http://SERVER_IP_ADDRESS:5000
```

To find the server's IP address, run `ipconfig` in a terminal and look for the IPv4 address.

---

## Starting / Stopping Manually

**Start:**
```bash
python serve.py
```

**Stop:**
```bash
taskkill /F /IM python.exe
```

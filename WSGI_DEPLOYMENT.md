# GPON Project - WSGI Deployment Guide

## Overview
This Flask application is now configured to run on a WSGI server. For **Windows**, it uses **Waitress** (pure Python, fully compatible). For **Linux/Production**, it can use **Gunicorn** with Eventlet for WebSocket support.

## Prerequisites
- Python 3.8+
- pip (Python package manager)
- Google Chrome (for Selenium automation)

## Installation

1. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

2. **Windows-specific setup:** No additional system dependencies needed (Waitress is pure Python)

3. **Linux/macOS setup:**
   - Install Chrome/Chromium for Selenium
   - Ubuntu/Debian:
     ```bash
     apt-get update
     apt-get install chromium-browser
     ```

## Running the Application

### Quick Start (Windows)

**Option 1: Using the batch script (Easiest)**
```bash
run_app.bat
```
This will automatically install dependencies and start the server.

**Option 2: Manual startup**
```bash
python -m waitress --port=5000 --host=0.0.0.0 wsgi:application
```

### Quick Start (Linux/macOS)

**Option 1: Using the shell script**
```bash
chmod +x run_app.sh
./run_app.sh
```

**Option 2: Manual startup**
```bash
python -m waitress --port=5000 --host=0.0.0.0 wsgi:application
```

### Development Server
```bash
python app.py
```
**Note:** Flask's development server should only be used for testing.

## Configuration

### Waitress Settings (Windows/Cross-platform)

**Command-line options:**
```bash
python -m waitress --port=5000 --host=0.0.0.0 --threads=8 wsgi:application
```

**Configuration options:**
- `--port=PORT` - Port number (default: 8080)
- `--host=HOST` - Host IP (0.0.0.0 for all interfaces)
- `--threads=N` - Number of worker threads (default: 4)
- `--asyncore-loop-timeout=SECONDS` - Socket poll timeout

**For production Waitress deployment, create a `waitress_config.py`:**
```python
from waitress import serve
from wsgi import application

if __name__ == '__main__':
    serve(
        application,
        host='0.0.0.0',
        port=5000,
        threads=8,
        max_request_body_size=16777216,  # 16MB
        ident='GPON',
        _quiet=False
    )
```

Then run:
```bash
python waitress_config.py
```

### Gunicorn Settings (Linux Production)

Edit `gunicorn_config.py` to customize:
- **Port:** Change `bind = "0.0.0.0:5000"`
- **Workers:** Adjust `workers = 4` (recommended: 2 * CPU_cores + 1)
- **Worker Class:** Use `eventlet` for WebSocket support (required for Flask-SocketIO)
- **Logging:** Configure `accesslog` and `errorlog` paths

## Running Behind Nginx (Recommended Setup)

### 1. Install Nginx
```bash
# Linux (Ubuntu/Debian)
sudo apt-get install nginx

# macOS
brew install nginx
```

### 2. Configure Nginx
Create `/etc/nginx/sites-available/gpon`:
```nginx
upstream gpon_app {
    server 127.0.0.1:5000;
}

server {
    listen 80;
    server_name your-domain.com;
    client_max_body_size 16M;

    location / {
        proxy_pass http://gpon_app;
        proxy_http_version 1.1;
        proxy_set_header Upgrade $http_upgrade;
        proxy_set_header Connection "upgrade";
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto $scheme;
    }

    # WebSocket support
    location /socket.io {
        proxy_pass http://gpon_app/socket.io;
        proxy_http_version 1.1;
        proxy_buffering off;
        proxy_set_header Upgrade $http_upgrade;
        proxy_set_header Connection "Upgrade";
        proxy_set_header Host $host;
        proxy_set_header X-Real-IP $remote_addr;
        proxy_set_header X-Forwarded-For $proxy_add_x_forwarded_for;
    }
}
```

### 3. Enable the site
```bash
sudo ln -s /etc/nginx/sites-available/gpon /etc/nginx/sites-enabled/
sudo systemctl restart nginx
```

## Running as a Systemd Service (Linux)

Create `/etc/systemd/system/gpon.service`:
```ini
[Unit]
Description=GPON Application
After=network.target

[Service]
Type=notify
User=www-data
WorkingDirectory=/path/to/GPON_PROJECT
ExecStart=/usr/bin/gunicorn --config gunicorn_config.py wsgi:app
Restart=on-failure
RestartSec=10

[Install]
WantedBy=multi-user.target
```

Enable and start the service:
```bash
sudo systemctl daemon-reload
sudo systemctl enable gpon
sudo systemctl start gpon
```

View logs:
```bash
sudo journalctl -u gpon -f
```

## WSGI Application Structure

### Entry Point: `wsgi.py`
- Imports the Flask app from `app.py`
- Exports the `application` variable for WSGI servers
- Compatible with Waitress, Gunicorn, uWSGI, etc.

### Main Application: `app.py`
- Flask application with SocketIO support
- No longer runs development server in production mode
- Works with all WSGI servers

### Startup Scripts
- **`run_app.bat`** - Windows batch script (auto-installs dependencies)
- **`run_app.sh`** - Linux/macOS shell script

### Requirements
- **Flask 2.3.3** - Web framework
- **Flask-SocketIO 5.3.6** - WebSocket support
- **Waitress 2.1.2** - WSGI server (Windows-compatible)
- **Eventlet 0.33.3** - Async support for WebSockets
- **Other dependencies** - Selenium, Paramiko, Pandas, etc.

## Troubleshooting

### Port Already in Use

**Windows:**
```powershell
# Find process using port 5000
netstat -ano | findstr :5000
# Kill the process
taskkill /PID <PID> /F
```

**Linux/macOS:**
```bash
# Find process using port 5000
lsof -i :5000
# Kill the process
kill -9 <PID>
```

### WebSocket Connection Issues
- Check browser console for connection errors
- Verify firewall allows WebSocket traffic (port 5000)
- Ensure `async_mode='threading'` in Flask-SocketIO
- Try connecting to `http://localhost:5000` first to test basic connectivity

## Performance Optimization

1. **Enable Caching:**
   - Configure browser caching in Nginx
   - Use CDN for static files

2. **Database Connection Pooling:**
   - If using a database, enable connection pooling

3. **Load Balancing:**
   - Run multiple Gunicorn instances on different ports
   - Use Nginx to load balance between them

4. **Monitoring:**
   - Use tools like New Relic, Datadog, or Prometheus
   - Monitor CPU, memory, and request latency

## Security Considerations

1. **Change Secret Key:** Update `app.secret_key` in app.py
2. **Use HTTPS:** Configure SSL/TLS certificates
3. **Rate Limiting:** Implement rate limiting for API endpoints
4. **CORS:** Review and restrict CORS settings
5. **Input Validation:** Ensure file uploads are validated
6. **Log Security:** Secure access to application logs

## Rollback Procedure

```bash
# Stop the service
sudo systemctl stop gpon

# Revert to previous version
git checkout <previous-commit>

# Reinstall dependencies if changed
pip install -r requirements.txt

# Start the service
sudo systemctl start gpon
```

## Additional Resources

- [Gunicorn Documentation](https://docs.gunicorn.org/)
- [Flask-SocketIO Documentation](https://flask-socketio.readthedocs.io/)
- [Nginx Reverse Proxy Guide](https://docs.nginx.com/nginx/admin-guide/web-server/reverse-proxy/)
- [WSGI Specification](https://www.python.org/dev/peps/pep-3333/)

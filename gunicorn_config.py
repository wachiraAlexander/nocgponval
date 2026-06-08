"""Gunicorn configuration file."""

# Server socket
bind = "0.0.0.0:5000"
backlog = 2048

# Worker processes
workers = 4  # (2 * CPU_cores) + 1
worker_class = "eventlet"  # Use eventlet for WebSocket support
worker_connections = 1000
timeout = 30
keepalive = 2

# Logging
accesslog = "-"
errorlog = "-"
loglevel = "info"
access_log_format = '%(h)s %(l)s %(u)s %(t)s "%(r)s" %(s)s %(b)s "%(f)s" "%(a)s"'

# Process naming
proc_name = "gpon_app"

# Server mechanics
daemon = False
pidfile = None
umask = 0
user = None
group = None
tmp_upload_dir = None

# SSL (optional - uncomment if using HTTPS)
# keyfile = "/path/to/keyfile"
# certfile = "/path/to/certfile"
# ssl_version = "TLSv1_2"

# Server hooks
def post_fork(server, worker):
    pass

def pre_fork(server, worker):
    pass

def pre_exec(server):
    pass

def when_ready(server):
    pass

# gunicorn_config.py
import multiprocessing

# Timeout configuration
timeout = 120  # 2 minutes (increase from default 30s)
graceful_timeout = 120
keepalive = 5

# Worker configuration
workers = 2  # Use 2 workers on free tier
worker_class = 'sync'
worker_connections = 1000

# Logging
accesslog = '-'
errorlog = '-'
loglevel = 'info'

# Memory management
max_requests = 100  # Restart workers after 100 requests to prevent memory leaks
max_requests_jitter = 10
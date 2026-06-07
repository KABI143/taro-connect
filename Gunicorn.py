# Gunicorn config for Render free plan
workers   = 1        # Free plan — 1 worker enough
timeout   = 120      # 2 min timeout (default 30s too short for Excel processing)
keepalive = 5
worker_class = "sync"
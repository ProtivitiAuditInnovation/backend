#!/bin/sh
# Optional explicit start command for Azure App Service (Linux, Python):
#   Configuration -> General settings -> Startup Command:  bash startup.sh
# Not required: Azure's default command also works because gunicorn reads gunicorn.conf.py.
exec gunicorn --bind=0.0.0.0:${PORT:-8000} --timeout 600 -c gunicorn.conf.py app:app

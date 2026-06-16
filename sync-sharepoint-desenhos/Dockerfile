FROM python:3.12-slim

WORKDIR /app

COPY sync-sharepoint-desenhos/requirements.txt /app/requirements.txt
RUN pip install --no-cache-dir -r /app/requirements.txt

COPY sync-sharepoint-desenhos/sync_recent_drawings.py /app/sync_recent_drawings.py

ENTRYPOINT ["python", "-u", "/app/sync_recent_drawings.py"]

FROM python:3.12-slim-bookworm

ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1 \
    ORACLE_CLIENT_LIB_DIR=/opt/oracle/instantclient \
    ORACLE_CLIENT_CONFIG_FILE=/opt/oracle/instantclient \
    LD_LIBRARY_PATH=/opt/oracle/instantclient

WORKDIR /app

RUN apt-get update \
    && apt-get install -y --no-install-recommends \
        ca-certificates \
        libaio1 \
        libnsl2 \
    && rm -rf /var/lib/apt/lists/*

COPY instantclient/instantclient_23_26 /opt/oracle/instantclient

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY oracle_to_db_material_rv.py .

ENTRYPOINT ["python", "oracle_to_db_material_rv.py"]

FROM python:3.11-slim

ENV PYTHONUNBUFFERED=1 \
    PORT=80 \
    PIP_NO_CACHE_DIR=1

WORKDIR /app

RUN pip config set global.index-url https://mirrors.cloud.tencent.com/pypi/simple && \
    pip config set global.trusted-host mirrors.cloud.tencent.com

COPY requirements.txt ./
RUN pip install --upgrade pip && pip install -r requirements.txt

COPY . .

EXPOSE 80
CMD ["sh", "-c", "uvicorn main:app --host 0.0.0.0 --port ${PORT:-80} --workers 1"]

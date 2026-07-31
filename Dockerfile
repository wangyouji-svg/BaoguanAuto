FROM python:3.12-slim

ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1

WORKDIR /app

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY scripts/backend_server.py ./scripts/backend_server.py
COPY 报关资料模板（只有一个商品）.xlsx ./
COPY 报关资料模板（有多个商品）.xlsx ./

RUN mkdir -p /app/scripts/generated

WORKDIR /app/scripts

EXPOSE 5000

CMD ["gunicorn", "-w", "1", "--threads", "4", "-b", "0.0.0.0:5000", "backend_server:app"]

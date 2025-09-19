FROM python:3.11-slim

WORKDIR /app

# instalar dependências básicas do sistema
RUN apt-get update && apt-get install -y --no-install-recommends \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

EXPOSE 8501

CMD ["streamlit", "run", "app_v2.py", "--server.port=8501", "--server.headless=true"]

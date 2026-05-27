FROM python:3.11-slim
WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY . .
RUN mkdir -p /app/data
EXPOSE 8000
CMD ["sh", "-c", "python migrate.py && uvicorn main:app --host 0.0.0.0 --port ${PORT:-8000}"]

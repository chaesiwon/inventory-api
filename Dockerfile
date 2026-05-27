FROM python:3.11-slim

WORKDIR /app

# 의존성 설치
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 앱 복사
COPY . .

# 데이터 디렉토리 (볼륨 마운트 포인트)
RUN mkdir -p /app/data

# 포트 (환경변수로 재정의 가능)
EXPOSE 8000

# DB 마이그레이션 후 서버 시작
CMD sh -c "python migrate.py && uvicorn main:app --host 0.0.0.0 --port ${PORT:-8000}"

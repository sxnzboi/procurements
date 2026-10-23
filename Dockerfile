FROM python:3.11-alpine

WORKDIR /app

# Pure Python — ไม่ต้องติดตั้ง dependency ภายนอกเพิ่มเติม
COPY . /app

EXPOSE 8080

ENV PORT=8080

CMD ["python3", "server.py"]

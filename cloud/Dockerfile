FROM python:3.13-slim

WORKDIR /app

# 依赖
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 应用代码
COPY server.py .
COPY public/ public/
COPY template/ template/

# 健康检查友好：云托管会请求 PORT
EXPOSE 3000

CMD ["python", "server.py"]

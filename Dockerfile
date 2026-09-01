# 水闸设计系统 - Fly.io 部署镜像
FROM python:3.12-slim

WORKDIR /app

# 先复制依赖文件，利用 Docker 层缓存
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 复制项目代码（排除项见 .dockerignore）
COPY . .

# 数据目录（挂载持久卷）
RUN mkdir -p /data

# 数据目录与种子库（云端首次启动自动恢复用户数据）
ENV DATA_DIR=/data
ENV SEED_DB=/app/users.db.seed

EXPOSE 8080

CMD ["sh", "-c", "gunicorn server:app --bind 0.0.0.0:8080 --workers 2 --timeout 180"]

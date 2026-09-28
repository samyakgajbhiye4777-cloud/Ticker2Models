FROM python:3.12-slim
WORKDIR /app
COPY . .
RUN pip install --no-cache-dir -r requirements.txt gunicorn
CMD gunicorn -b 0.0.0.0:${PORT:-8000} -w 1 --threads 4 --timeout 180 app:app

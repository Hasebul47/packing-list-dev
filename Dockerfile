# ---------------------------------------------------------
# PO -> Packing List - multi-buyer Streamlit app  (port 8003) 
# ---------------------------------------------------------
FROM python:3.12-slim

ENV PYTHONUNBUFFERED=1 \
  PYTHONDONTWRITEBYTECODE=1 \
  PIP_NO_CACHE_DIR=1 \
  PIP_DISABLE_PIP_VERSION_CHECK=1 \
  PORT=8003 \
  STREAMLIT_SERVER_HEADLESS=true \
  STREAMLIT_SERVER_ADDRESS=0.0.0.0 \
  STREAMLIT_SERVER_PORT=8003 \
  STREAMLIT_BROWSER_GATHER_USAGE_STATS=false

# curl is only used by the health check
RUN apt-get update && apt-get install -y --no-install-recommends curl \
  && rm -rf /var/lib/apt/lists/*

# Non-root user
RUN useradd -m -u 1000 appuser
WORKDIR /app

# Python deps first (better layer caching)
COPY requirements.txt /app/requirements.txt
RUN pip install -r /app/requirements.txt

# Project code (+ the Excel templates the buyers need - see .dockerignore)
COPY . /app
RUN chown -R appuser:appuser /app
USER appuser

EXPOSE 8003

HEALTHCHECK --interval=30s --timeout=10s --start-period=25s --retries=3 \
  CMD curl -fs http://localhost:8003/_stcore/health || exit 1

CMD ["streamlit", "run", "app.py", "--server.port=8003", "--server.address=0.0.0.0"]

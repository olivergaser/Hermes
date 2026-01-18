FROM python:3.9-slim-bookworm

# Prevent Python from writing pyc files to disc
ENV PYTHONDONTWRITEBYTECODE=1
# Prevent Python from buffering stdout and stderr
ENV PYTHONUNBUFFERED=1

# Install system dependencies
# WeasyPrint needs Pango, Cairo, etc.
# LibreOffice is required for Office document conversion.
# Poppler-utils is required for PDF to TIFF conversion.
RUN apt-get update && apt-get install -y --no-install-recommends \
    build-essential \
    libffi-dev \
    libcairo2 \
    libpango-1.0-0 \
    libpangocairo-1.0-0 \
    libgdk-pixbuf2.0-0 \
    shared-mime-info \
    libreoffice-writer \
    libreoffice-calc \
    libreoffice-impress \
    poppler-utils \
    curl \
    ca-certificates \
    && rm -rf /var/lib/apt/lists/*

# Install uv for fast dependency management
COPY --from=ghcr.io/astral-sh/uv:latest /uv /bin/uv

WORKDIR /app

# Copy dependency definitions
COPY pyproject.toml uv.lock README.md ./

# Install dependencies into the system python environment
RUN uv pip install --system .

# Copy the rest of the application
COPY . .

# Set entrypoint
ENTRYPOINT ["python", "converter.py"]
# Default arguments (can be overridden)
CMD ["--input", "./eml", "--output", "./eml", "--format", "pdf"]

FROM python:3.11-slim

# Evita criação de arquivos .pyc e força stdout imediato
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1
ENV DEBIAN_FRONTEND=noninteractive

# Instala bibliotecas de sistema para suporte a GUI (Tkinter, X11) e bibliotecas de PDF
RUN apt-get update && apt-get install -y --no-install-recommends \
    python3-tk \
    tk-dev \
    tcl-dev \
    libx11-6 \
    libxext6 \
    libxrender1 \
    libxinerama1 \
    libxi6 \
    libxrandr2 \
    libxcursor1 \
    build-essential \
    && rm -rf /var/lib/apt/lists/*

WORKDIR /app

# Copia requirements e instala dependências com cache eficiente
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# Copia os arquivos da aplicação
COPY . .

# Comando padrão executa a suíte completa de testes automatizados com pytest
CMD ["pytest", "tests/", "-v"]

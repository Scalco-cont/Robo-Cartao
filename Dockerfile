FROM python:3.11-slim

WORKDIR /app

# Instalar dependências necessárias para compilar algumas bibliotecas Python se necessário
RUN apt-get update && apt-get install -y \
    gcc \
    && rm -rf /var/lib/apt/lists/*

COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

COPY . .

# Cria os diretórios necessários
RUN mkdir -p uploads outputs

# Expõe a porta original
EXPOSE 2929

# Usar Gunicorn com --preload para eliminar cold start (aquece tudo antes do primeiro usuario)
CMD ["gunicorn", "--bind", "0.0.0.0:2929", "--workers", "4", "--timeout", "120", "--preload", "app:app"]

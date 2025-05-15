# Usando uma imagem Python SLIM baseada no Debian Bookworm
FROM python:3.12-slim-bookworm

ENV PYTHONUNBUFFERED 1 # Garante que os logs do Python apareçam imediatamente
ENV VENV_PATH=/opt/venv
ENV PATH=${VENV_PATH}/bin:${PATH}

# Define o diretório de trabalho principal da aplicação
WORKDIR /app

# Instalar dependências de sistema:
# - build-essential: Pode ser necessário para compilar algumas extensões Python.
# - Dependências para WeasyPrint: Pango, Cairo, GDK-PixBuf são as principais.
# - fonts-liberation: Boas fontes fallback que funcionam bem com WeasyPrint.
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
    build-essential \
    # Dependências para WeasyPrint e suas sub-dependências
    libpango-1.0-0 \
    libpangocairo-1.0-0 \
    libcairo2 \
    libgdk-pixbuf-2.0-0 \
    libffi-dev \          
    shared-mime-info \    
    fonts-liberation && \
    rm -rf /var/lib/apt/lists/*

# Criar o Ambiente Virtual
RUN python3 -m venv ${VENV_PATH}

# Atualizar pip dentro do ambiente virtual
RUN pip install --upgrade pip

# Copiar o arquivo de dependências primeiro para aproveitar o cache do Docker
COPY ./src/requirements.txt /app/requirements.txt

# Instalar as dependências Python
RUN pip install --no-cache-dir -r requirements.txt

# Copiar o restante da sua pasta src (que contém app.py, templates/, utils/, etc.) para /app
COPY ./src/ /app/


# Expõe a porta que o Flask estará rodando (se não definido no docker-compose)
EXPOSE 8000 

# Comando para executar a aplicação
# app.py deve estar agora em /app/app.py
CMD ["python3", "app.py"]

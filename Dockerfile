# Usando a imagem Python completa baseada no Debian Bookworm
FROM python:3.12-bookworm

# Variáveis de ambiente para Rust e o ambiente virtual
# ENV RUSTUP_HOME=/usr/local/rustup \
#     CARGO_HOME=/usr/local/cargo \
#     VENV_PATH=/opt/venv
ENV VENV_PATH=/opt/venv

# Adiciona os diretórios bin do Cargo (Rust) e do venv ao PATH.
# O venv vem depois no PATH para que seus executáveis (python, pip) tenham precedência.
# ENV PATH=${CARGO_HOME}/bin:${VENV_PATH}/bin:${PATH}
ENV PATH=${VENV_PATH}/bin:${PATH}

WORKDIR /app

# Instalar dependências de sistema:
# - build-essential: Para compilar extensões C/C++ (ex: numpy, pandas)
# - curl: Para baixar o instalador do Rust
# (Outras bibliotecas -dev podem ser adicionadas aqui se necessário no futuro)
# Instalar dependências de sistema:
RUN apt-get update && \
    apt-get install -y --no-install-recommends \
    build-essential \
    libreoffice \
    xvfb \
    fonts-liberation \
    unoconv \
    python3-uno && \
    rm -rf /var/lib/apt/lists/*

# Instalar Rust (usando o PATH já configurado com CARGO_HOME)
# RUN curl --proto '=https' --tlsv1.2 -sSf https://sh.rustup.rs | sh -s -- --default-toolchain stable -y

# Criar o Ambiente Virtual (usando a variável VENV_PATH)
RUN python3 -m venv ${VENV_PATH}

# Atualizar pip dentro do ambiente virtual
# (o pip do venv será usado devido à configuração do PATH)
RUN pip install --upgrade pip

# Copiar o arquivo de requisitos
COPY ./src/requirements.txt ./

# Instalar as dependências Python do requirements.txt
# --no-cache-dir é uma boa prática para manter as camadas menores, mesmo que o tamanho total não seja o problema principal
RUN pip install --no-cache-dir -r requirements.txt

# Copiar o código da sua aplicação (ex: pasta src)
WORKDIR /src

COPY ./src/ /src/

CMD ["python3", "app.py"]
# ------------------------------------------------------------------------------------


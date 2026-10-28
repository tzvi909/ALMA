FROM python:3.12-slim
ENV DEBIAN_FRONTEND=noninteractive
RUN apt-get update && apt-get install -y --no-install-recommends libreoffice-calc-nogui libreoffice-script-provider-python python3-uno fonts-dejavu-core && rm -rf /var/lib/apt/lists/*

# --- Stage 1: Builder ---
FROM golang:1.24-alpine AS builder

# 1. Install dependencies
RUN apk add --no-cache build-base wget pkgconf

WORKDIR /tmp

# 2. Download libxls 1.6.2
RUN wget https://github.com/libxls/libxls/releases/download/v1.6.2/libxls-1.6.2.tar.gz \
    && tar -xzvf libxls-1.6.2.tar.gz

# 3. PATCH & Compile libxls (CRITICAL STEP)
WORKDIR /tmp/libxls-1.6.2

# Penjelasan Patch:
# 1. Patch XlsReader.h: Menambah include <string> & <cstdint> untuk C++ compiler baru.
# 2. Patch xlstypes.h: Mendefinisikan manual tipe BYTE, WORD, DWORD yang hilang di Alpine Linux.
RUN sed -i '1i #include <cstdint>\n#include <string>' cplusplus/XlsReader.h \
    && sed -i '1i #include <stdint.h>\ntypedef uint8_t BYTE;\ntypedef uint16_t WORD;\ntypedef uint32_t DWORD;' include/libxls/xlstypes.h \
    && ./configure --prefix=/usr \
    && make \
    && make install

# 4. Setup Aplikasi Go
WORKDIR /app
COPY go.mod go.sum ./
RUN go mod download

COPY . .

# 5. Build Go App
# Kita set environment variable agar CGO tahu di mana library berada
ENV CGO_CFLAGS="-I/usr/include/libxls -I/usr/include"
ENV CGO_LDFLAGS="-L/usr/lib -lxlsreader"

RUN CGO_ENABLED=1 GOOS=linux go build -ldflags="-s -w" -o excel-service main.go

# --- Stage 2: Runner ---
FROM alpine:latest

# Install dependencies runtime
RUN apk add --no-cache libstdc++ libgcc

WORKDIR /root/

# Copy Binary
COPY --from=builder /app/excel-service .

# Copy Library libxls (.so files)
COPY --from=builder /usr/lib/libxlsreader.so* /usr/lib/
COPY --from=builder /usr/lib/libxls.so* /usr/lib/

RUN mkdir uploads

EXPOSE 3020

CMD ["./excel-service"]

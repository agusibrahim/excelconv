# --- Stage 1: Builder ---
# UBAH DISINI: Menggunakan versi 1.24 sesuai requirement go.mod Anda
FROM golang:1.24-alpine AS builder

WORKDIR /app

COPY go.mod go.sum ./
RUN go mod download

COPY . .

RUN CGO_ENABLED=0 GOOS=linux go build -o excel-service main.go

# --- Stage 2: Runner ---
FROM alpine:latest

RUN apk --no-cache add ca-certificates
WORKDIR /root/

COPY --from=builder /app/excel-service .
RUN mkdir uploads

EXPOSE 3020

CMD ["./excel-service"]

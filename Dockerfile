FROM node:20-alpine

# Set timezone ke Asia/Jakarta (WIB) untuk timestamp laporan yang tepat
RUN apk add --no-cache tzdata
ENV TZ=Asia/Jakarta

WORKDIR /app

# Copy dependency definition & install
COPY package*.json ./
RUN npm install --omit=dev

# Copy bot code
COPY index.js ./
COPY ecosystem.config.js ./

# Jalankan bot
CMD ["node", "index.js"]

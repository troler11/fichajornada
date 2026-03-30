# ==========================================
# ESTÁGIO 1: Build (Compilação do TypeScript)
# ==========================================
FROM node:20-alpine AS builder

# Define o diretório de trabalho
WORKDIR /usr/src/app

# Copia os arquivos de configuração de pacotes
COPY package*.json ./
COPY tsconfig.json ./

# Instala TODAS as dependências (incluindo as de desenvolvimento, como o 'typescript')
RUN npm install

# Copia todo o código fonte da pasta src
COPY src/ ./src/

# Roda o compilador do TypeScript (gera a pasta /dist)
RUN npm run build

# ==========================================
# ESTÁGIO 2: Produção (Imagem final leve)
# ==========================================
FROM node:20-alpine

WORKDIR /usr/src/app

COPY package*.json ./
RUN npm install --only=production

# Copia o código compilado
COPY --from=builder /usr/src/app/dist ./dist

# ADICIONE ESTA LINHA: Copia a tela do sistema (o HTML) para o servidor
COPY public/ ./public/

EXPOSE 3000
CMD ["npm", "start"]

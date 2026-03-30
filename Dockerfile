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

# Copia apenas o package.json
COPY package*.json ./

# Instala APENAS as dependências de produção (ignora o typescript e @types)
RUN npm install --only=production

# Copia a pasta /dist compilada que foi gerada no ESTÁGIO 1
COPY --from=builder /usr/src/app/dist ./dist

# Expõe a porta do Express
EXPOSE 3000

# Inicia o servidor rodando o arquivo JavaScript final
CMD ["npm", "start"]

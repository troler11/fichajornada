# Usa uma imagem oficial do Node.js, versão leve (Alpine) para economizar espaço no servidor
FROM node:20-alpine

# Define o diretório de trabalho dentro do contêiner
WORKDIR /usr/src/app

# Copia apenas os arquivos de dependência primeiro (melhora o cache do Docker)
COPY package*.json ./

# Instala as dependências do projeto
RUN npm install

# Copia o restante do código da aplicação para o contêiner
COPY . .

# (Opcional) Se estiver usando TypeScript, descomente a linha abaixo para compilar o código
# RUN npm run build

# Expõe a porta que a sua aplicação web vai usar (ex: 3000)
EXPOSE 3000

# Comando para iniciar o servidor
CMD ["npm", "start"]
# Se for usar o arquivo compilado do TypeScript, mude para: CMD ["node", "dist/index.js"]

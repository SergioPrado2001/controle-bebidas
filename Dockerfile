# Imagem base
FROM node:18

# Diretório dentro do container
WORKDIR /app

# Copia arquivos
COPY package*.json ./
RUN npm install

COPY . .

# Porta do app
EXPOSE 3000

# Comando de inicialização
CMD ["npm", "start"]


FROM node:20

RUN apt-get update && apt-get install -y python3 python3-pip

WORKDIR /app

COPY . .

RUN npm install
RUN pip3 install pandas openpyxl
EXPOSE 3000
CMD ["node", "server.js"]

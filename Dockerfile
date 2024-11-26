# Node.js 이미지를 기반으로 합니다.
FROM node:18

# 작업 디렉토리를 설정합니다.
WORKDIR /usr/src/app

# package.json과 package-lock.json을 복사합니다.
COPY package*.json ./

# 종속성을 설치합니다.
RUN npm install

# 애플리케이션 코드를 복사합니다.
COPY . .

# 애플리케이션이 실행될 포트를 설정합니다.
EXPOSE 3333

# 컨테이너가 실행될 때 실행할 명령어를 설정합니다.
CMD ["node", "index.js"]
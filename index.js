require('dotenv').config();
const express = require('express');
const AWS = require('aws-sdk');
const app = express();
const port = 3333; // 포트 번호를 3333번으로 설정

// AWS 설정
AWS.config.update({
  region: 'ap-northeast-2', // 서울 리전
});

const dynamodb = new AWS.DynamoDB.DocumentClient();

// DynamoDB 연결 확인
const params = { TableName: 'UserSelectedDocuments' };
dynamodb.scan(params, (err, data) => {
  if (err) {
    console.error('DynamoDB 연결 오류:', err);
  } else {
    console.log('DynamoDB 연결 성공:', data);
  }
});

// 라우터 설정
const excelRouter = require('./excel.js');
app.use('/', excelRouter);

app.listen(port, () => {
  console.log(`Server is running on port ${port}`);
});

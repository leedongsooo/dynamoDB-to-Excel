

const express = require('express');
const router = express.Router();
const AWS = require('aws-sdk');
const ExcelJS = require('exceljs');
const cors = require('cors');
const path = require('path');

// AWS 설정
AWS.config.update({
  region: 'ap-northeast-2',
  accessKeyId: process.env.AWS_ACCESS_KEY_ID,
  secretAccessKey: process.env.AWS_SECRET_ACCESS_KEY,
});

const dynamodb = new AWS.DynamoDB.DocumentClient();

router.use(cors());

// 문자열 정규화 함수
function normalizeString(str) {
  if (!str) return '';
  return str.normalize('NFC');
}

// ISMS ID 비교 함수
function compareISMSIds(a, b) {
  if (!a || !b) return 0;
  
  const partsA = a.toString().split('.').map(Number);
  const partsB = b.toString().split('.').map(Number);

  for (let i = 0; i < Math.max(partsA.length, partsB.length); i++) {
    const numA = partsA[i] || 0;
    const numB = partsB[i] || 0;
    if (numA !== numB) {
      return numA - numB;
    }
  }
  return 0;
}

// 행 높이 계산 함수
function calculateRowHeight(text) {
  if (!text) return 0;
  const lines = text.split('\n');
  const averageCharsPerLine = 50; // 한 줄에 들어갈 수 있는 평균 글자 수
  
  let totalLines = 0;
  lines.forEach(line => {
    totalLines += Math.ceil(line.length / averageCharsPerLine);
  });
  
  return totalLines;
}

// reason 필드들 추출 함수
function extractReasons(doc) {
  const reasons = [];
  let i = 1;
  
  while (doc[`reason${i}`]) {
    const reason = doc[`reason${i}`].trim();
    if (reason.toLowerCase() !== 'none') {
      reasons.push(reason);
    }
    i++;
  }
  
  return reasons;
}

// ISMS 항목별 데이터 그룹화 함수
function groupByISMSItem(userDocs, evidenceDocs) {
  const groupedData = new Map();

  // 정책 문서 그룹화
  userDocs?.forEach((doc) => {
    if (!doc.ISMSID) return;
    
    const ismsId = doc.ISMSID.trim();
    if (!groupedData.has(ismsId)) {
      groupedData.set(ismsId, {
        ismsId: ismsId,
        contents: new Set(),
        reasons: new Map(),
        policies: new Set(),
        evidences: new Set(),
      });
    }
    
    if (doc.Content && doc.Content.trim().toLowerCase() !== 'none') {
      groupedData.get(ismsId).contents.add(doc.Content.trim());
    }
    
    if (doc.full_path && doc.full_path.trim().toLowerCase() !== 'none') {
      groupedData.get(ismsId).policies.add(doc.full_path.trim());
    }
  });

  // 증적 문서 그룹화
  evidenceDocs?.forEach((doc) => {
    if (!doc.ISMSItem) return;
    
    const ismsItem = doc.ISMSItem.trim();
    if (!groupedData.has(ismsItem)) {
      groupedData.set(ismsItem, {
        ismsId: ismsItem,
        contents: new Set(),
        reasons: new Map(),
        policies: new Set(),
        evidences: new Set(),
      });
    }

    if (doc.FileName && doc.FileName.trim().toLowerCase() !== 'none') {
      groupedData.get(ismsItem).evidences.add(doc.FileName.trim());
    }

    const reasons = extractReasons(doc);
    if (reasons.length > 0 && doc.FileName && doc.FileName.trim().toLowerCase() !== 'none') {
      groupedData.get(ismsItem).reasons.set(doc.FileName.trim(), reasons);
    }
  });

  return Array.from(groupedData.values())
    .map(item => ({
      ...item,
      contents: Array.from(item.contents),
      reasons: Array.from(item.reasons.entries()),
      policies: Array.from(item.policies),
      evidences: Array.from(item.evidences)
    }))
    .sort((a, b) => compareISMSIds(a.ismsId, b.ismsId));
}

// 엑셀 템플릿에 데이터 매핑 함수
function mapDataToTemplate(sheet, data) {
  const MIN_ROW = 3;
  const DEFAULT_ROW_HEIGHT = 50;   
  const MIN_CONTENT_HEIGHT = 70;   
  const MAX_ROW_HEIGHT = 1000;      
  const HEIGHT_PER_LINE = 15;      

  let currentRow = MIN_ROW;

  try {
    while (currentRow <= sheet.rowCount) {
      const ismsCell = sheet.getCell(`F${currentRow}`);
      const ismsId = ismsCell.value?.toString().trim();
      
      const iCell = sheet.getCell(`I${currentRow}`);
      const jCell = sheet.getCell(`J${currentRow}`);
      const kCell = sheet.getCell(`K${currentRow}`);
      
      if (ismsId) {
        const matchingData = data.find(item => 
          item.ismsId && compareISMSIds(item.ismsId.trim(), ismsId) === 0
        );

        if (matchingData) {
          // I열에 Content와 Reasons 작성
          let contentText = matchingData.contents
            .filter(Boolean)
            .filter(content => content.toLowerCase() !== 'none')
            .map(content => normalizeString(content))
            .join('\n');

          // Reasons 추가 (들여쓰기와 Rich Text 포함)
          const reasonsText = matchingData.reasons
            .map(([filename, reasons]) => {
              if (reasons && reasons.length > 0) {
                const validReasons = reasons.filter(reason => 
                  reason && reason.toLowerCase() !== 'none'
                );
                if (validReasons.length > 0) {
                  return {
                    text: `  ${normalizeString(filename)}:\n${validReasons.map(reason => 
                      `    - ${normalizeString(reason)}`).join('\n')}`,
                    filename: filename
                  };
                }
              }
              return null;
            })
            .filter(Boolean);

          if (contentText && reasonsText.length > 0) {
            // Rich Text를 사용하여 Content와 Reasons를 결합
            iCell.value = {
              richText: [
                { text: contentText + '\n' },
                ...reasonsText.flatMap((item, index) => [
                  { 
                    text: '  ' + normalizeString(item.filename) + ':\n',
                    font: { bold: true }
                  },
                  { 
                    text: item.text.split('\n')
                      .slice(1) // 첫 번째 줄(파일명)을 제외
                      .join('\n') + (index < reasonsText.length - 1 ? '\n' : '')
                  }
                ])
              ]
            };
          } else if (reasonsText.length > 0) {
            iCell.value = {
              richText: reasonsText.flatMap((item, index) => [
                { 
                  text: '  ' + normalizeString(item.filename) + ':\n',
                  font: { bold: true }
                },
                { 
                  text: item.text.split('\n')
                    .slice(1)
                    .join('\n') + (index < reasonsText.length - 1 ? '\n' : '')
                }
              ])
            };
          } else {
            iCell.value = contentText;
          }
          
          applyCellStyle(iCell);

          // J열에 정책 현황 작성
          const policyText = matchingData.policies
            .filter(Boolean)
            .filter(policy => policy.toLowerCase() !== 'none')
            .map(policy => normalizeString(policy))
            .join('\n');
          
          jCell.value = policyText;
          applyCellStyle(jCell);

          // K열에 증적 현황 작성
          const evidenceText = matchingData.evidences
            .filter(Boolean)
            .filter(evidence => evidence.toLowerCase() !== 'none')
            .map(evidence => normalizeString(evidence))
            .join('\n');
          
          kCell.value = evidenceText;
          applyCellStyle(kCell);
          
          // 행 높이 계산
          const iCellLines = calculateRowHeight(typeof iCell.value === 'object' ? 
            iCell.value.richText.map(rt => rt.text).join('') : 
            iCell.value);
          const jCellLines = calculateRowHeight(policyText);
          const kCellLines = calculateRowHeight(evidenceText);
          
          const maxLines = Math.max(iCellLines, jCellLines, kCellLines);
          
          if (maxLines > 0) {
            const calculatedHeight = Math.max(
              MIN_CONTENT_HEIGHT,
              maxLines * HEIGHT_PER_LINE
            );
            sheet.getRow(currentRow).height = Math.min(calculatedHeight, MAX_ROW_HEIGHT);
          } else {
            sheet.getRow(currentRow).height = DEFAULT_ROW_HEIGHT;
          }
        } else {
          sheet.getRow(currentRow).height = DEFAULT_ROW_HEIGHT;
        }
      } else {
        sheet.getRow(currentRow).height = DEFAULT_ROW_HEIGHT;
      }

      currentRow++;
    }
  } catch (error) {
    console.error(`Error in mapDataToTemplate at row ${currentRow}:`, error);
    throw error;
  }
}

// 셀 스타일 적용 함수
function applyCellStyle(cell) {
  cell.alignment = {
    vertical: 'middle',
    horizontal: 'left',
    wrapText: true
  };
  cell.font = {
    name: '맑은 고딕',
    size: 9
  };
  cell.border = {
    top: { style: 'thin' },
    left: { style: 'thin' },
    bottom: { style: 'thin' },
    right: { style: 'thin' }
  };
}

// Excel 다운로드 API 엔드포인트
router.get('/api/download-excel', async (req, res) => {
  let workbook;
  
  try {
    const [userDocsData, evidenceData] = await Promise.all([
      dynamodb.scan({ 
        TableName: 'UserSelectedDocuments',
        ConsistentRead: true 
      }).promise(),
      dynamodb.scan({ 
        TableName: 'Evidence_Metadata',
        ConsistentRead: true 
      }).promise(),
    ]);

    if (!userDocsData.Items?.length && !evidenceData.Items?.length) {
      return res.status(404).json({ 
        error: '데이터가 없습니다.',
        message: 'DynamoDB 테이블에서 데이터를 찾을 수 없습니다.' 
      });
    }

    workbook = new ExcelJS.Workbook();
    const templatePath = path.join(__dirname, './template/PIM template.xlsx');
    await workbook.xlsx.readFile(templatePath);
    
        // 첫 번째 시트를 활성화
        workbook.views = [{
          x: 0, y: 0, width: 30000, height: 20000,
          firstSheet: 0, activeTab: 0, visibility: 'visible'
        }];
    

    const groupedData = groupByISMSItem(
      userDocsData.Items || [], 
      evidenceData.Items || []
    );

    const sheet1 = workbook.getWorksheet('1.관리체계 수립 및 운영');
    const sheet2 = workbook.getWorksheet('2.보호대책 요구사항');

    if (!sheet1 || !sheet2) {
      throw new Error('Required worksheets not found in template');
    }

    await Promise.all([
      mapDataToTemplate(sheet1, groupedData.filter(item => !item.ismsId.startsWith('2.'))),
      mapDataToTemplate(sheet2, groupedData.filter(item => item.ismsId.startsWith('2.')))
    ]);

    const fileName = encodeURIComponent('ISMS_Status.xlsx')
      .replace(/['()]/g, escape)
      .replace(/\*/g, '%2A')
      .replace(/%(?:7C|60|5E)/g, encodeURIComponent);

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.setHeader(
      'Content-Disposition',
      `attachment; filename=${fileName}; filename*=UTF-8''${fileName}`
    );
    res.setHeader('Cache-Control', 'no-cache');
    
    await workbook.xlsx.write(res);
    res.end();

  } catch (error) {
    console.error('Excel 생성 중 오류:', error);
    
    if (!res.headersSent) {
      res.status(500).json({
        error: '파일 생성 중 오류가 발생했습니다.',
        message: error.message,
        stack: process.env.NODE_ENV === 'development' ? error.stack : undefined
      });
    }
    
  } finally {
    if (workbook) {
      workbook = null;
    }
  }
});

module.exports = router;

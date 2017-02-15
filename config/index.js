export default {
  excel: {
    SheetName: 'シート1',
    Columns: {
      0: 'sentenceExample',
      1: 'systemResponses',
      2: 'type',
    },
    HeadRowNumber: 0,
    InputExcelFilePath: __dirname + '/../xlsx/elastic-search.xlsx',
    OutputJsonFilePath: __dirname + '/../json/data.json'
  },
};
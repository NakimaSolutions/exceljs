const path = require('path');
const Excel = require('../lib/exceljs.nodejs');
const HrStopwatch = require('./utils/hr-stopwatch');

const filename = process.argv[2];

const wb = new Excel.stream.xlsx.WorkbookWriter({
  filename: `${filename}`,
  useStyles: true,
});
const ws1 = wb.addWorksheet('Fooo1', {views: [{showGridLines: false}]});

const imageId1 = wb.addImage({
  filename:  path.join(__dirname, 'data/bubbles.jpg'),
  extension: 'jpg',
});
ws1.addImage(imageId1, {
  tl: {col: 0.25, row: 0.7},
  ext: {width: 160, height: 60},
});

const ws2 = wb.addWorksheet('Fooo2');
const imageId2 = wb.addImage({
  filename: path.join(__dirname, 'data/image2.png'),
  extension: 'png',
});
ws2.addImage(imageId2, {
  tl: {col: 3.5, row: 0.25},
  br: {col: 4, row: 2},
});

const stopwatch = new HrStopwatch();
stopwatch.start();

wb.commit()
  .then(() => {
    const micros = stopwatch.microseconds;
    console.log('Done.');
    console.log('Time taken:', micros);
  })
  .catch(error => {
    console.log(error.message);
  });

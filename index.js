const 
  xlsx = require('xlsx'),
  yml = require('require-yml'),
  mysql = require('mysql');

const 
  config = yml('config'), 
  pool = mysql.createPool(config.mysql),
  excelData = xlsx.readFile(config.excel).Sheets.Sheet1;



// 转换excel数据为二维数组
let excelArr = [];
for(let i in excelData) {
  if(i[0] === '!') continue;
  let row = +i.replace(/[^\d]/g, '') - 1;
  let col = +i.replace(/\d/g, '').toUpperCase().charCodeAt() - 65;
  if(!excelArr[row]) excelArr[row] = [];
  excelArr[row][col] = excelData[i].v;
}

let table = excelArr.shift()[0], heads = excelArr.shift().join(','), values = excelArr.shift(), len = excelArr.length + 1, i = 1;

const insertData = (conn, values) => conn.query(`INSERT INTO ${table} (${heads}) VALUES (${values.map(v => `'${v}'`).join(',')})`, (e, f) => {
  if(e) return console.log(e);
  console.log(`导入成功 ${Math.floor(i++/len*100)}%`)
  values = excelArr.shift();
  if(values) insertData(conn, values);
  else console.log('导入完毕!')
});

// 循环插入数据
pool.getConnection((err, conn) => {

  if(err) return console.log(err);

  if(values) insertData(conn, values);
});
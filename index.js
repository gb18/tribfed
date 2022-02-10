const https = require('https');
var xlsx = require('xlsx');
const { Pool, Client } = require('pg');

const pool = new Pool({
    user: 'postgres',
    host: 'localhost',
    database: 'tribfed',
    password: 'admin',
    port: 5432
});

convertXlsx();

function convertXlsx() {
    var obj = xlsx.readFile('C:/Users/giubu/Downloads/tributos federais.xlsx');
    var col1;
    var col2;
    var col3;
    var col4;

    // pool.query("INSERT INTO tribfed(cidade, ano, impostos, valor) values ('cidade', 2020, 'teste', 25.00)")
    var sheetnames = obj.SheetNames;
    sheetnames.forEach(async function(y) {
        var worksheet = obj.Sheets[y];
        var headers = {};
        var data = [];
        for(z in worksheet) {
            if(z[0] === '!') continue;
            var tt = 0;
            
            for (var i = 0; i < z.length; i++) {
                if (!isNaN(z[i])) {
                    tt = i;
                    break;
                }
            };
            
            var col = z.substring(0,tt);
            var row = parseInt(z.substring(tt));
            var value = worksheet[z].v;

            if(row == 1 && value) {
                headers[col] = value;
                continue;
            }

            if(col == 'A') {col1 = value}
            if(col == 'B') {col2 = value}
            if(col == 'C') {col3 = value}
            if(col == 'D') {
                col4 = value
                await pool.query("INSERT INTO tribfed(cidade, ano, impostos, valor) VALUES (' "+ col1 + 
                            "'," + col2 + ",'" + col3 + "'," + col4 + ")");
                console.log(col1, col2, col3, col4)
                col1 = '';
                col2 = 0;
                col3 = '';
                col4 = 0;
            }

            
    
            if(!data[row]) data[row]={};
            data[row][headers[col]] = value;
        }

        data.shift();
        data.shift();
    });
}
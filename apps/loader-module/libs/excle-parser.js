const XLSX = require('xlsx');

function extractDataFromExcel(fileBuffer, rules) {
    const workbook = XLSX.read(fileBuffer);
    const sheetNames = workbook.SheetNames;

    if (!Array.isArray(rules[0])) rules = [rules];

    const knex = require('knex')({
        client: 'sqlite3',
        connection: {
            // filename: ':memory:',
            filename: `./apps/loader-module/api/app-${Math.floor(Math.random() * 100)}.db`,
        },
    });

    rules.forEach((rulesForSheet, i) => {
        const forSheet = rulesForSheet.filter(rule => 'sheet' in rule)[0]?.['sheet'];

        const sheetName = forSheet ? forSheet : sheetNames[i];
        const worksheet = workbook.Sheets[sheetName];

        let sheetObj = XLSX.utils.sheet_to_json(worksheet, { raw: false, header: 1, blankrows: true });

        let headers = [];
        let dataTypes = [];
        let data = [];

        // usecols
        const usecols = rulesForSheet.filter(rule => 'usecols' in rule)[0]?.['usecols'];
        if (usecols) {
            const range = parseRange(usecols);
            sheetObj.forEach((row, i) => {
                sheetObj[i] = selectByRange(row, range);
            });
        }

        rulesForSheet.forEach(rule => {
            const ruleName = Object.keys(rule)[0];

            // take header
            if (ruleName === 'take' && rule[ruleName] === 'headers') {
                headers = sheetObj[0];
                dataTypes = sheetObj[1];
                sheetObj = sheetObj.slice(1);
            }

            // take data
            if (ruleName === 'take' && rule[ruleName] === 'data') data = sheetObj;

            // skip row
            if (ruleName === 'skip') sheetObj = sheetObj.slice(rule[ruleName]);
        });

        let dataObj = [];

        data.forEach(row => {
            let obj = {};

            row.forEach((el, i) => {
                obj[headers[i]] = el;
            });

            dataObj.push(obj);
        });

        saveToSqlite(knex, sheetName, dataObj);
    });
}

function parseRange(str) {
    const result = [];

    const parts = str.split(',');
    for (let i = 0; i < parts.length; i++) {
        if (parts[i].indexOf('-') > 0) {
            const range = parts[i].split('-');
            const start = parseInt(range[0]);
            const end = parseInt(range[1]);
            for (let j = start; j <= end; j++) {
                result.push(j);
            }
        } else {
            result.push(parseInt(parts[i]));
        }
    }

    return result;
}

function selectByRange(arr, range) {
    return arr.filter((_, index) => range.includes(index + 1));
}

async function saveToSqlite(knex, tableName, data) {
    const colNames = Object.keys(data[0]);

    try {
        await knex.schema
            .createTable(tableName, table => {
                table.increments('id');

                colNames.forEach(colName => {
                    // TODO
                    table.string(colName);
                });
            });

        await knex(tableName).insert(data);
    } catch (err) {
        console.error(err);
    };
}

module.exports = extractDataFromExcel;

const XLSX = require('xlsx');

function extractDataFromExcel(fileBuffer, rules) {
    const workbook = XLSX.read(fileBuffer);
    const sheetNames = workbook.SheetNames;

    if (!Array.isArray(rules[0])) rules = [rules];

    let excelData = {};

    rules.forEach((rulesForSheet, i) => {
        const forSheet = rulesForSheet.filter(rule => 'sheet' in rule)[0]?.['sheet'];

        const sheetName = forSheet || sheetNames[i];
        const worksheet = workbook.Sheets[sheetName];

        let sheetObj = XLSX.utils.sheet_to_json(worksheet, { raw: true, header: 1, blankrows: true });

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

        excelData[sheetName] = dataObj;
    });

    // return excelData;

    let success = true;
    let workbookErr = [];

    rules.forEach((rulesForSheet, i) => {
        const forSheet = rulesForSheet.filter(rule => 'sheet' in rule)[0]?.['sheet'];

        const sheetName = forSheet || sheetNames[i];

        const data = excelData[sheetName];

        let sheetErrObj = {};
        sheetErrObj[sheetName] = []

        data.forEach((row, rowIndex) => {
            const keys = Object.keys(row);

            let rawEvalStr = '';

            keys.forEach((key, i) => {
                const value = typeof row[key] == 'string' ? `'${row[key]}'` : row[key];
                rawEvalStr += `const f${i + 1} = ${value};\n`;
            });

            // if (rowIndex == 3) console.log(rawEvalStr);

            const expRules = rulesForSheet.filter(rule => rule['exp']);
            // const checkRes = expRules.map(rule => eval(rawEvalStr + rule['exp']));
            
            let rowErr = [];
            expRules.forEach(rule => {
                if (!eval(rawEvalStr + rule['exp'])) {
                    rowErr.push(rule['exp']);
                    success = false;
                }
            });

            if (rowErr.length !== 0) sheetErrObj[sheetName].push({ row: rowIndex + 1, failed: rowErr });
        });

        workbookErr.push(sheetErrObj);
    });

    if (success) {
        return { success, data: excelData };
    } else {
        return { success, data: workbookErr };
    }
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

module.exports = extractDataFromExcel;

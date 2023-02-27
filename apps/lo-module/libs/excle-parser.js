const XLSX = require('xlsx');


async function extractDataFromExcel(rules, fileBuffer) {
    const knex = require('knex')({
        client: 'sqlite3',
        connection: {
            filename: './app.db',
        },
    });

    const workbook = XLSX.read(fileBuffer, { cellDates: true });
    const sheetNames = workbook.SheetNames;

    if (!Array.isArray(rules[0])) rules = [rules];

    for (let i = 0; i < rules.length; i++) {
        const rulesForSheet = rules[i];
        const sheetName = rulesForSheet.filter(rule => 'sheet' in rule)[0]?.['sheet'] || sheetNames[i];
        const worksheet = workbook.Sheets[sheetName];

        let sheetArr = XLSX.utils.sheet_to_json(worksheet, { header: 1, blankrows: true });

        let headers = [];
        let dataTypes = [];
        let data = [];
        let selectedData = [];

        const usecols = rulesForSheet.filter(rule => 'usecols' in rule)[0]?.['usecols'];
        if (usecols) {
            const range = parseRange(usecols);
            sheetArr.forEach((row, i) => {
                sheetArr[i] = selectByRange(row, range);
            });
        }

        for (let j = 0; j < rulesForSheet.length; j++) {
            const rule = rulesForSheet[j];
            const ruleName = Object.keys(rule)[0];

            // take header
            if (ruleName === 'take' && rule[ruleName] === 'headers') {
                headers = sheetArr[0];
                dataTypes = sheetArr[1];
                sheetArr = sheetArr.slice(1);
            }

            // take data
            if (ruleName === 'take' && rule[ruleName] === 'data') data = sheetArr;

            // skip row
            if (ruleName === 'skip') sheetArr = sheetArr.slice(rule[ruleName]);
        }

        data.forEach(row => {
            let nr = {};
            headers.forEach((header, index) => {
                nr[header] = row[index];
            });
            selectedData.push(nr);
        });

        // ================================================================

        let transformedData = [];

        for (let j = 0; j < selectedData.length; j++) {
            let row = selectedData[j];
            
            const transformRules = rulesForSheet.filter(rule => rule['rule'] && !rule['onnew']);
            let newRow = {};
            let excludeCols = [];

            for (let k = 0; k < transformRules.length; k++) {
                const ruleV = transformRules[k].rule.replaceAll(' ', '');

                if (ruleV.match(/[\w]+/g)[1] === ruleV.split('=')[1]) {
                    newRow[ruleV.split('=')[0]] = row[ruleV.split('=')[1]];

                    if (!excludeCols.includes(ruleV.split('=')[1])) excludeCols.push(ruleV.split('=')[1]);
                }

                if (ruleV.includes('split(')) {
                    const args = ruleV.match(/(?<=\().+?(?=\))/g)[0]
                        .replaceAll('"', '')
                        .replaceAll("'", '')
                        .split(',');
                    const index = parseInt(ruleV.match(/[\d]/g)[0]);

                    newRow[ruleV.split('=')[0]] = row[args[0]].split(args[1])[index];

                    if (!excludeCols.includes(args[0])) excludeCols.push(args[0]);
                }

                if (ruleV.includes('+')) {
                    const operands = ruleV.split('=')[1].split('+');
                    newRow[ruleV.split('=')[0]] = '';

                    operands.forEach(operand => {
                        newRow[ruleV.split('=')[0]] += /['"]/.test(operand) ? operand.replaceAll("'", '') : row[operand];

                        if (!/['"]/.test(operand) && !excludeCols.includes(operand)) excludeCols.push(operand);
                    });
                }
            }


            const replaceRules = rulesForSheet.filter(rule => rule['rule'] && rule['onnew']);

            for (let k = 0; k < replaceRules.length; k++) {
                const ruleV = replaceRules[k].rule.replaceAll(' ', '');
                const args = [...ruleV.matchAll(/[\w.]+/g)].map(el => el.flat()[0]);

                const newColName = args[0];
                const fromTable = args[1].split('.')[0];
                const fromTableField = args[1].split('.')[1];
                const whereField = args[2];
                const colName = args[3];

                const selectQuery = await knex(fromTable).select(fromTableField).where(whereField, row[colName]);

                if (selectQuery.length === 1) {
                    newRow[newColName] = selectQuery[0][fromTableField];
                    if (!excludeCols.includes(colName)) excludeCols.push(colName);
                }

                if (selectQuery.length === 0 && replaceRules[k].onnew === 'create') {
                    let _tempObj = {};
                    _tempObj[whereField] = row[colName];
                    const _tempSqlRes = await knex(fromTable).insert(_tempObj, ['*']);

                    newRow[newColName] = _tempSqlRes[0][fromTableField];
                    if (!excludeCols.includes(colName)) excludeCols.push(colName);
                }

                if (selectQuery.length === 0 && replaceRules[k].onnew === 'ignore') {
                    newRow[newColName] = null;
                    if (!excludeCols.includes(colName)) excludeCols.push(colName);
                }

                if (selectQuery.length === 0 && replaceRules[k].onnew === 'fail') {
                    const errMsg = `[row: ${j + 1}] The following rule threw an exception: ${ruleV}; ${colName}:${row[colName]}`;
                    console.log(errMsg);
                    return { success: false, message: errMsg };
                }
            }

            excludeCols.forEach(colName => {
                delete row[colName];
            });
            newRow = { ...row, ...newRow };
            transformedData.push(newRow);
        }

        // ================================================================

        let expResForRowArr = [];
        let sheetCheckPassed = true;

        transformedData.forEach((row, rowIndex) => {
            let evalVars = '';
            Object.keys(row).forEach(key => {
                const value = typeof row[key] == 'string' ? `'${row[key]}'` : row[key];
                evalVars += `const ${key} = ${value};\n`;
            });

            const expRules = rulesForSheet.filter(rule => rule['exp']);

            let expResForRow = [];
            let rowCheckPassed = true;

            for (let j = 0; j < expRules.length; j++) {
                const ruleV = expRules[j].exp;
                const ruleRes = eval(evalVars + ruleV);

                if (!ruleRes) rowCheckPassed = false;

                expResForRow.push({ ...expRules[j], result: ruleRes });
            }

            if (!rowCheckPassed) sheetCheckPassed = false;

            expResForRowArr.push({ row: rowIndex + 1, rules: expResForRow, passed: rowCheckPassed });
        });

        if (!sheetCheckPassed) {
            return {
                sheetName,
                success: false,
                rows: expResForRowArr,
            }
        }

        // Loading to DB
        const colNames = Object.keys(transformedData[0]);

        await knex.schema.createTable(sheetName, table => {
            colNames.forEach(colName => {
                let colDataType = '';
    
                switch (typeof transformedData[0][colName]) {
                    case 'number':
                        if (colName === 'id') {
                            colDataType = 'increments';
                        } else {
                            colDataType = Number.isInteger(transformedData[0][colName]) ? 'integer' : 'float';
                        }
                        break;
                    default:
                        colDataType = 'string';
                        break;
                }
    
                table[colDataType](colName);
            });
        });

        await knex(sheetName).insert(transformedData);

        // ================================================================

        // console.log(sheetArr);
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

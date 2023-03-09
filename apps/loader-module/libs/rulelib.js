const fs = require('node:fs');
const XLSX = require('xlsx');


async function extractDataFromExcel(rules, fileBuffer) {
    if (fs.existsSync('./app.db')) fs.unlinkSync('./app.db');

    const knex = require('knex')({
        client: 'sqlite3',
        connection: {
            filename: './app.db',
        },
        useNullAsDefault: true,
    });

    const workbook = XLSX.read(fileBuffer, { cellDates: true });
    const sheetNames = workbook.SheetNames;

    let sheetNameIndex = 0;

    if (!Array.isArray(rules[0])) rules = [rules];

    for (let i = 0; i < rules.length; i++) {
        const rulesForSheet = rules[i];
        const sheetName = rulesForSheet.filter(rule => 'sheet' in rule)[0]?.['sheet'] || sheetNames[sheetNameIndex];
        const worksheet = workbook.Sheets[sheetName];

        const selectRules = rulesForSheet.filter(rule => ('take' in rule) || ('skip' in rule) || ('usecols' in rule));
        sheetNameIndex += selectRules.length > 0 ? 1 : 0;
        let data = extractData(selectRules, worksheet);


        const nonSelectRules = rulesForSheet.filter(rule => !('take' in rule) && !('skip' in rule) && !('usecols' in rule));
        for (let j = 0; j < nonSelectRules.length; j++) {
            try {
                const rule = nonSelectRules[j];

                // Преобразование имен полей
                if (
                    !(Array.isArray(rule))
                    && !('onnew' in rule)
                    && !Object.keys(rule)[0].includes('.')
                    && !('table' in rule)
                    && !('join' in rule)
                    && !('fields' in rule)
                    && !('to_table' in rule)
                    && !('from_table' in rule)
                    && !('key_fields' in rule)
                    && !('onduplicate' in rule)
                ) {
                    const newColNames = Object.keys(rule);
                    let excludeCols = [];

                    data.forEach((row, rowIndex) => {
                        let newRow = {};

                        newColNames.forEach(newColName => {
                            if (rule[newColName] === rule[newColName].match(/[\w]+/g)?.[0]) {
                                newRow[newColName] = row[rule[newColName]];

                                if (!excludeCols.includes(rule[newColName])) excludeCols.push(rule[newColName]);
                            }

                            if (/(?<=split\().+?(?=\))/g.test(rule[newColName])) {
                                const splitArgs = rule[newColName].match(/(?<=split\().+?(?=\))/g)[0].replaceAll(' ', '').split(',');
                                splitArgs[1] = eval(splitArgs[1]);
                                splitArgs[2] = eval(splitArgs[2]);
                                newRow[newColName] = row[splitArgs[0]].split(splitArgs[1])[splitArgs[2]];

                                if (!excludeCols.includes(splitArgs[0])) excludeCols.push(splitArgs[0]);
                            }

                            if (rule[newColName].includes('+')) {
                                const operands = rule[newColName].replaceAll(' ', '').split('+');
                                newRow[newColName] = '';

                                operands.forEach(operand => {
                                    newRow[newColName] += /['"]/.test(operand) ? eval(operand) : row[operand];

                                    if (!/['"]/.test(operand) && !excludeCols.includes(operand)) excludeCols.push(operand);
                                });
                            }
                        });

                        excludeCols.forEach(colName => delete row[colName]);
                        data[rowIndex] = { ...row, ...newRow };
                    });
                }

                // Проверка типов и ограничений
                if (Array.isArray(rule) && rule.length > 0 && typeof rule[0] === 'string') {
                    let expResForRowArr = [];

                    data.forEach((row, rowIndex) => {
                        let evalVars = '';
                        Object.keys(row).forEach(key => {
                            const value = typeof row[key] == 'string' ? `'${row[key]}'` : row[key];
                            evalVars += `const ${key} = ${value};\n`;
                        });

                        let failedRules = [];

                        rule.forEach(checkRule => {
                            const ruleRes = eval(evalVars + checkRule);
                            if (!ruleRes) failedRules.push(checkRule);
                        });

                        if (failedRules.length > 0) expResForRowArr.push({ row: rowIndex + 1, failedRules });
                    });

                    if (expResForRowArr.length > 0) {
                        // TODO
                        return {
                            sheetName,
                            failedRows: expResForRowArr,
                        }
                    }
                }

                // Замена значение-код / код-значение
                if (!(Array.isArray(rule)) && 'onnew' in rule) {
                    const newColName = Object.keys(rule).filter(key => key !== 'onnew')[0];

                    const args = rule[newColName].match(/[\w]+/g);

                    for (let rowIndex = 0; rowIndex < data.length; rowIndex++) {
                        let row = data[rowIndex];

                        const query = await knex(args[0]).select(args[1]).where(args[2], row[args[3]]);

                        if (query.length === 1) {
                            row[newColName] = query[0][args[1]];
                            delete row[args[3]];
                        }

                        if (query.length === 0 && rule.onnew === 'create') {
                            const createSql = await knex(args[0]).insert({ [args[2]]: row[args[3]] }, ['*']);
                            row[newColName] = createSql[0][args[1]];
                            delete row[args[3]];
                        }

                        if (query.length === 0 && rule.onnew === 'fail') {
                            throw new Error(`[row: ${i + 1}] The following rule threw an exception: ${rule[newColName]}; ${args[3]}:${row[args[3]]}`);
                        }

                        if (query.length === 0 && rule.onnew === 'ignore') {
                            newRow[newColName] = null;
                            delete row[args[3]];
                        }
                    }
                }

                // Разбивка таблиц
                if (!(Array.isArray(rule)) && !('onnew' in rule) && Object.keys(rule)[0].includes('.')) {
                    const toTC = Object.keys(rule)[0].split('.');
                    const fromTC = rule[Object.keys(rule)[0]].split('.');

                    const query = await knex(fromTC[0]).select(fromTC[1]);

                    if (await knex.schema.hasTable(toTC[0])) {
                        const q = await knex(toTC[0]).select('*');

                        await knex.schema.table(toTC[0], table => {
                            let colDataType = '';

                            switch (typeof query[0][fromTC[1]]) {
                                case 'number':
                                    if (toTC[1] === 'id') {
                                        colDataType = 'increments';
                                    } else {
                                        colDataType = Number.isInteger(q[0][fromTC[1]]) ? 'integer' : 'float';
                                    }
                                    break;
                                default:
                                    colDataType = 'string';
                                    break;
                            }

                            table[colDataType](toTC[1]);
                        });

                        for (let k = 0; k < q.length; k++) {
                            let whereQ = [];
                            Object.keys(q[k]).forEach(field => whereQ.push(`${field}='${q[k][field]}'`));

                            const nc = Object.keys(query[k])[0]
                            await knex.raw(`UPDATE ${toTC[0]} SET ${nc}='${query[k][nc]}' WHERE ${whereQ.join(' AND ')}`);
                        }
                    } else {
                        await knex.schema.createTable(toTC[0], table => {
                            let colDataType = '';

                            switch (typeof query[0][fromTC[1]]) {
                                case 'number':
                                    if (toTC[1] === 'id') {
                                        colDataType = 'increments';
                                    } else {
                                        colDataType = Number.isInteger(q[0][fromTC[1]]) ? 'integer' : 'float';
                                    }
                                    break;
                                default:
                                    colDataType = 'string';
                                    break;
                            }

                            table[colDataType](toTC[1]);
                        });

                        await knex(toTC[0]).insert(query);
                    }
                }

                // Слияние таблиц
                if (!(Array.isArray(rule)) && 'table' in rule && 'join' in rule && 'fields' in rule) {
                    const fields = Object.keys(rule.fields).map(fieldName => `${rule.fields[fieldName]} as ${fieldName}`);
                    const tables = rule.join.replaceAll(' ', '').split('=').map(item => item.split('.')[0]);

                    const query = await knex.raw(`SELECT ${fields.join()} FROM ${tables[0]} INNER JOIN ${tables[1]} ON ${rule.join};`);

                    if (query.length > 0) {
                        const colNames = Object.keys(query[0]);
                        await knex.schema.createTable(rule.table, table => {
                            colNames.forEach(colName => {
                                let colDataType = '';

                                switch (typeof query[0][colName]) {
                                    case 'number':
                                        if (colName === 'id') {
                                            colDataType = 'increments';
                                        } else {
                                            colDataType = Number.isInteger(query[0][colName]) ? 'integer' : 'float';
                                        }
                                        break;
                                    default:
                                        colDataType = 'string';
                                        break;
                                }

                                table[colDataType](colName);
                            });
                        });

                        await knex(rule.table).insert(query);
                    }
                }

                // Тестовая заливка данных
                if (!(Array.isArray(rule)) && 'to_table' in rule && 'from_table' in rule && 'key_fields' in rule) {
                    if (rule.clear === true) {
                        await knex.schema.dropTable(rule.to_table);

                        const query = await knex(rule.from_table).select('*');

                        await knex.schema.createTable(rule.to_table, table => {
                            Object.keys(query[0]).forEach(colName => {
                                let colDataType = '';

                                switch (typeof query[0][colName]) {
                                    case 'number':
                                        if (colName === 'id') {
                                            colDataType = 'increments';
                                        } else {
                                            colDataType = Number.isInteger(query[0][colName]) ? 'integer' : 'float';
                                        }
                                        break;
                                    default:
                                        colDataType = 'string';
                                        break;
                                }

                                table[colDataType](colName);
                            });
                        });

                        await knex(rule.to_table).insert(query);
                    } else {
                        // onduplicate: "fail"
                        if (rule.onduplicate === 'fail') {
                            const byFileds = rule.key_fields.replaceAll(' ', '').split(',');
                            const intersections = await knex(rule.to_table).select(byFileds)
                                .intersect(knex(rule.from_table).select(byFileds));
                            
                            if (intersections.length > 0) {
                                throw new Error(`Error when trying to load the ${rule.from_table} table into the ${rule.to_table} table`);
                            }

                            const q = await knex(rule.from_table).select('*');
                            await knex(rule.to_table).insert(q);
                        }

                        // onduplicate: "new-over-old"
                        if (rule.onduplicate === 'new-over-old') {
                            const colNames = Object.keys(await knex(rule.to_table).columnInfo());
                            const coalesce = colNames.map(colName => `coalesce(${rule.to_table}.${colName}, ${rule.from_table}.${colName}) as ${colName}`).join();
                            const onFields = rule.key_fields.replaceAll(' ', '').split(',').map(field => `${rule.to_table}.${field}=${rule.from_table}.${field}`).join(' AND ');

                            const q = await knex.raw(`SELECT ${coalesce} FROM ${rule.to_table} LEFT JOIN ${rule.from_table} ON ${onFields} ` +
                                'UNION ' +
                                `SELECT ${coalesce} FROM ${rule.from_table} LEFT JOIN ${rule.to_table} ON ${onFields};`);

                            await knex(rule.to_table).insert(q).onConflict(...rule.key_fields.replaceAll(' ', '').split(',')).merge('*');
                        }

                        // onduplicate: "old-over-new"
                        if (rule.onduplicate === 'old-over-new') {
                            const colNames = Object.keys(await knex(rule.to_table).columnInfo());
                            const coalesce = colNames.map(colName => `coalesce(${rule.from_table}.${colName}, ${rule.to_table}.${colName}) as ${colName}`).join();
                            const onFields = rule.key_fields.replaceAll(' ', '').split(',').map(field => `${rule.to_table}.${field}=${rule.from_table}.${field}`).join(' AND ');

                            const q = await knex.raw(`SELECT ${coalesce} FROM ${rule.to_table} LEFT JOIN ${rule.from_table} ON ${onFields} ` +
                                'UNION ' +
                                `SELECT ${coalesce} FROM ${rule.from_table} LEFT JOIN ${rule.to_table} ON ${onFields};`);

                            await knex(rule.to_table).insert(q).onConflict(...rule.key_fields.replaceAll(' ', '').split(',')).merge('*');
                        }

                        // onduplicate: "old"
                        if (rule.onduplicate === 'old') {
                            const q = await knex(rule.from_table).select('*');
                            await knex(rule.to_table).insert(q).onConflict(...rule.key_fields.replaceAll(' ', '').split(',')).ignore();
                        }

                        // onduplicate: "new"
                        if (rule.onduplicate === 'new') {
                            const q = await knex(rule.from_table).select('*');
                            await knex(rule.to_table).insert(q).onConflict(...rule.key_fields.replaceAll(' ', '').split(',')).merge('*');
                        }
                    }
                }
            } catch (err) {
                console.log(err);
                return { error: err.message };
            }
        }

        // Load to DB
        if (data.length > 0) {
            const colNames = Object.keys(data[0]);
            await knex.schema.createTable(sheetName, table => {
                colNames.forEach(colName => {
                    let colDataType = '';

                    switch (typeof data[0][colName]) {
                        case 'number':
                            if (colName === 'id') {
                                colDataType = 'increments';
                            } else {
                                colDataType = Number.isInteger(data[0][colName]) ? 'integer' : 'float';
                            }
                            break;
                        default:
                            colDataType = 'string';
                            break;
                    }

                    table[colDataType](colName);
                });
            });

            await knex(sheetName).insert(data);
        }
    }
}

function extractData(rules, worksheet) {
    let sheetArr = XLSX.utils.sheet_to_json(worksheet, { header: 1, blankrows: true });

    let headers = [];
    let dataTypes = [];
    let data = [];

    let selectedData = [];

    // usecols
    const usecols = rules.filter(rule => 'usecols' in rule)[0]?.usecols;

    if (usecols) {
        const range = parseRange(usecols);
        sheetArr.forEach((row, i) => {
            sheetArr[i] = selectByRange(row, range);
        });
    }

    rules.forEach(rule => {
        // take header
        if (rule.take === 'headers') {
            headers = sheetArr[0];
            dataTypes = sheetArr[1];
            sheetArr = sheetArr.slice(1);
        }

        // take data
        if (rule.take === 'data') data = sheetArr;

        // skip row
        if (rule.skip) sheetArr = sheetArr.slice(rule.skip);
    });

    data.forEach(row => {
        let nr = {};
        headers.forEach((header, index) => {
            nr[header] = row[index];
        });
        selectedData.push(nr);
    });

    return selectedData;
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

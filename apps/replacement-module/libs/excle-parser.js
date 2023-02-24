const XLSX = require('xlsx');

const knex = require('knex')({
    client: 'sqlite3',
    connection: {
        filename: './dictionary.db',
    },
});

const moduleDB = require('knex')({
    client: 'sqlite3',
    connection: {
        filename: './apps/replacement-module/api/app.db',
    },
});

async function createTableIfNotExist() {
    try {
        if (!(await knex.schema.hasTable('X'))) {
            await knex.schema.createTable('X', table => {
                table.increments('id');
                table.string('name');
            });
        }
        if (!(await knex.schema.hasTable('Y'))) {
            await knex.schema.createTable('Y', table => {
                table.increments('id');
                table.string('codename');
            });
        }
    } catch (e) {
        console.error(e);
    }
}

async function extractDataFromExcel(fileBuffer, rules) {
    // TODO
    await createTableIfNotExist();

    const workbook = XLSX.read(fileBuffer);
    const sheetName = workbook.SheetNames[0];

    const worksheet = workbook.Sheets[sheetName];

    await moduleDB.schema.dropTableIfExists(workbook.SheetNames[0]);

    const sheetObj = XLSX.utils.sheet_to_json(worksheet, { raw: true });

    let newTableData = [];

    for (let i = 0; i < sheetObj.length; i++) {
        const row = sheetObj[i];
        let nr = {};
        nr[Object.keys(row)[0]] = row[Object.keys(row)[0]];

        for (let j = 0; j < rules.length; j++) {
            const rule = rules[j];
            const args = [...rule.matchAll(/[\w.]+/g)].map(el => el.flat()[0]);

            const newColName = args[0];
            const tableName = args[1].split('.')[0];
            const fieldName = args[1].split('.')[1];
            const findByField = args[2];
            const colName = args[3];
            const actionName = args[4];
            const actionValue = args[5];

            const selectQuery = await knex(tableName).select(fieldName).where(findByField, row[colName]);
            
            try {
                if (selectQuery.length === 1) {
                    nr[newColName] = selectQuery[0][fieldName];
                }

                if (selectQuery.length === 0 && actionName === 'onnew' && actionValue === 'fail')
                    throw `[row: ${i + 1}] The following rule threw an exception: ${rule}`;

                if (selectQuery.length === 0 && actionName === 'onnew' && actionValue === 'create') {
                    let t1 = {};
                    t1[findByField] = row[colName];
                    const t2 = await knex(tableName).insert(t1, ['*']);
                    nr[newColName] = t2[0][fieldName];
                }

                if (selectQuery.length === 0 && actionName === 'onnew' && actionValue === 'ignore') {
                    let t1 = {};
                    t1[findByField] = row[colName];
                    const t2 = await knex(tableName).insert(t1, ['*']);
                    nr[newColName] = t2[0][fieldName];
                }
            } catch (err ) {
                console.log(err);
                return { success: false, message: err };
            }
        }

        newTableData.push(nr);
    }

    if (!(await knex.schema.hasTable(workbook.SheetNames[0]))) { 
        const colNames = Object.keys(newTableData[0]);

        await moduleDB.schema.createTable(workbook.SheetNames[0], table => {
            colNames.forEach(cn => {
                let fieldType = '';

                switch (typeof newTableData[0][cn]) {
                    case 'number':
                        fieldType = Number.isInteger(newTableData[0][cn]) ? 'integer' : 'float';
                        break;
                    default:
                        fieldType = 'string';
                        break;
                }

                table[fieldType](cn);
            });
        });
    }
    
    for (let i = 0; i < newTableData.length; i++) {
        await moduleDB(workbook.SheetNames[0]).insert(newTableData[i]);
    }

    return { success: true };
}

module.exports = extractDataFromExcel;

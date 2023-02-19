const XLSX = require('xlsx');

function transformSheet(fileBuffer, rules) {
    const workbook = XLSX.read(fileBuffer);
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];

    const sheetJson = XLSX.utils.sheet_to_json(worksheet, { raw: false, blankrows: true });

    const newSheetJson = [];

    sheetJson.forEach(row => {
        let newRow = {};
        let excludeCols = [];

        rules.forEach(rule => {
            const rule_split = rule.replaceAll(' ', '').split('=');
            const destColName = rule_split[0];

            if (rule_split[1].match(/^[a-zA-Z][a-zA-Z0-9]*/)?.flat()[0] === rule_split[1]) {
                newRow[destColName] = row[rule_split[1]];
                
                if (!excludeCols.includes(rule_split[1])) excludeCols.push(rule_split[1]);
            }

            if (rule_split[1].match(/^split\(/)?.flat()[0] === 'split(') {
                const args = rule_split[1].match(/\([a-zA-Z]*.*[\)]/).flat()[0].slice(1, -1).split(',');
                const sourceColName = args[0];
                const separator = args[1].replaceAll("'", '');
                const index = rule_split[1].match(/[0-9][0-9]*/).flat()[0];

                newRow[destColName] = row[sourceColName].split(separator)[index];

                if (!excludeCols.includes(sourceColName)) excludeCols.push(sourceColName);
            }

            if (/[+]/.test(rule_split[1])) {
                const operands = rule_split[1].split('+');
                newRow[destColName] = ''
                
                operands.forEach(operand => {
                    newRow[destColName] += operand.includes("'") ? operand.replaceAll("'", '') : row[operand];
                    
                    if (!operand.includes("'") && !excludeCols.includes(operand)) excludeCols.push(operand);
                });
            }
        });

        let newRowObj = {};
        Object.keys(row).forEach(el => {
            if (!excludeCols.includes(el)) newRowObj[el] = row[el];
        });

        const transformedRow = {...newRowObj, ...newRow}
        newSheetJson.push(transformedRow);
    });

    return { data: newSheetJson, sheetName: workbook.SheetNames[0] };
}

module.exports = transformSheet;

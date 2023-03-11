const XLSX = require('xlsx');

function typesMapping(dataType) {
    dataType = dataType.toLowerCase();

    // https://www.prisma.io/docs/reference/api-reference/prisma-schema-reference#model-field-scalar-types
    const int_types = ['integer'];
    const float_types = ['float', 'double', 'm', 'm3', 'ton', 'c', 'bar', 'sstvd', 'tvd'];
    const decimal_types = ['usd'];
    const dateTime_type = ['date'];

    if (int_types.includes(dataType)) return 'Int';
    if (float_types.includes(dataType)) return 'Float';
    if (decimal_types.includes(dataType)) return 'Decimal';
    if (dateTime_type.includes(dataType)) return 'DateTime';

    return 'String';
}

function convertToPrisma(fileBuffer) {
    const workbook = XLSX.read(fileBuffer, { cellDates: true });
    const sheetNames = workbook.SheetNames;

    let schema = '';

    sheetNames.forEach(sheetName => {
        const worksheet = workbook.Sheets[sheetName];
        const sheetObj = XLSX.utils.sheet_to_json(worksheet, { header: 1, blankrows: true });

        if (sheetName === '_list') {
            sheetObj.join().split(',,').map(el => el.split(',')).forEach(arr => {
                schema += `enum ${arr[0]} {\n`;
                arr.slice(1).forEach(item => schema += `  ${item}\n`);
                schema += '}\n\n';
            });
        } else {
            const colNames = sheetObj[0];
            const colTypes = sheetObj[1];
            schema += `model ${sheetName} {\n`;

            for (let i = 0; i < colNames.length; i++) {
                const colType = typesMapping(colTypes[i]);
                schema += `  ${colNames[i]} ${colType}\n`;
            }

            schema += '}\n\n';
        }
    });

    return schema;
}

module.exports = convertToPrisma;

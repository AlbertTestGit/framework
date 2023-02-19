let input = document.getElementById('input');
let btn = document.getElementById('btn');

async function upload() {
    const formData = new FormData();
    const fileField = document.getElementById('file');

    formData.append('rules', input.value);
    formData.append('file', fileField.files[0]);

    try {
        const response = await fetch('./api', {
            method: 'POST',
            body: formData,
        });
        const result = await response.json();

        createFile(result.data, result.sheetName);

        console.log(JSON.stringify(result));
        alert('Успех!')
    } catch (error) {
        console.error(error);
        alert('Ошибка!')
    }
}

function createFile(data, sheetName) {
    const worksheet = XLSX.utils.json_to_sheet(data);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
    XLSX.writeFile(workbook, "data.xlsx");
}

btn.addEventListener('click', upload);

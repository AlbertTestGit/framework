let input = document.getElementById('input');
let btn = document.getElementById('btn');

async function upload() {
    const formData = new FormData();
    const fileField = document.getElementById('file');

    formData.append('rules', input.value);
    formData.append('file', fileField.files[0]);

    const response = await fetch('./api', {
        method: 'POST',
        body: formData,
    });
    const result = await response.json();

    // console.log(JSON.stringify(result));
    if (result.success) {
        alert('Успех!');
    } else {
        alert('Ошибка!');
        console.log(result.message);
    }
}

btn.addEventListener('click', upload);

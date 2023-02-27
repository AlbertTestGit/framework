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
        // const result = await response.json();
        // console.log(JSON.stringify(result));
        // alert('Успех!')
    } catch (error) {
        console.error(error);
        // alert('Ошибка!')
    }
}

btn.addEventListener('click', upload);

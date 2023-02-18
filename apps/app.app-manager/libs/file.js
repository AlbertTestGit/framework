/**
 * Load data of file as base64 string with mime (i.e. image/jpeg:base64,UUUU...)
 * @param {File} file file to load 
 * @returns {Promise<string>} 
 */
function readFileData(file) {
    return new Promise(resolve => {
        const reader = new FileReader();
        reader.addEventListener('load', (event) => {
            resolve({ name: file.name, data: event.target.result });
        });
        reader.readAsDataURL(file);
    })
}

/**
 * Read all files as {name, data: base64}
 * @param {HTMLNode} input input type=file with files 
 * @returns {Promise<[name:string, data:string]>} List with selected fiels data
 */
async function readFiles(input) {
    return Promise.all([...input.files].map(readFileData))
}

export {
    readFileData,
    readFiles
}
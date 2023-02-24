const express = require('express');
const multer = require('multer');

const extractDataFromExcel = require('../libs/excle-parser');


const router = express.Router();

router.use(express.json())
router.use(multer({storage: multer.memoryStorage()}).single('file'));

router.post('/', async (req, res) => {
    // TODO: валидировать входные данные
    const rules = req.body.rules.split('\r\n');
    const file = req.file;

    const result = await extractDataFromExcel(file.buffer, rules);

    if (!result.success) res.statusCode = 400;
    return res.send(result);
});

module.exports = router;

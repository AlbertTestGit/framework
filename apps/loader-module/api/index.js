const express = require('express');
const multer = require('multer');

const extractDataFromExcel = require('../libs/excle-parser');

const router = express.Router();

router.use(express.json())
router.use(multer({storage: multer.memoryStorage()}).single('file'));

router.post('/', async (req, res) => {
    // TODO: валидировать входные данные
    const rules = eval(req.body.rules);
    const file = req.file;
    
    extractDataFromExcel(file.buffer, rules);

    return res.send({rules, file: file.originalname});
});

module.exports = router;

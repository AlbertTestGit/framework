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

    // return res.send({ ok: true });
    return res.send(await extractDataFromExcel(rules, file.buffer));
});

module.exports = router;

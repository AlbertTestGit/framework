const express = require('express');
const multer = require('multer');

const transformSheet = require('../libs/excle-transform');

const router = express.Router();

router.use(express.json())
router.use(multer({storage: multer.memoryStorage()}).single('file'));

router.post('/', async (req, res) => {
    // TODO: валидировать входные данные
    const rules = req.body.rules.split('\r\n');
    const file = req.file;

    return res.send(transformSheet(file.buffer, rules));
});

module.exports = router;

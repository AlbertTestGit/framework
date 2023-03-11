const express = require('express');
const multer = require('multer');

const convertToPrisma = require('../libs/xlsx2prisma');

const router = express.Router();

router.use(express.json())
router.use(multer({storage: multer.memoryStorage()}).single('file'));

router.post('/', async (req, res) => {
    const file = req.file;

    return res.send({ data: convertToPrisma(file.buffer) });
});

module.exports = router;

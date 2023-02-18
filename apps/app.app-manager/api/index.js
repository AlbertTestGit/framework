const express = require('express')
const multer = require('multer')
const { readdirSync, rmSync } = require('fs')
const AdmZip = require('adm-zip')

const conf = require('./conf')
const router = express.Router()
const upload = multer({ dest: conf.apps_dir })

router.get('/', (req, res) => {
    let dirs = readdirSync(conf.apps_dir).filter(x => !x.endsWith('.js'))
    res.send(dirs)
})

// TODO: Restrictions
router.get('/:id/@export', (req, res) => {
    const zip = new AdmZip()
    zip.addLocalFolder(conf.apps_dir + '/' + req.params.id)

    const downloadName = `${req.params.id}.zip`;
    const data = zip.toBuffer();
    res.set('Content-Type', 'application/octet-stream');
    res.set('Content-Disposition', `attachment; filename=${downloadName}`);
    res.set('Content-Length', data.length);
    res.send(data);
})

// TODO: Restrictions
router.post('/', upload.single('file'), async (req, res) => {
    try {
        const raw_zip = req.body.data.substring(req.body.data.indexOf('base64,') + 7)
        const name = req.body.name.replace('.zip', '')
        const zip = new AdmZip(Buffer.from(raw_zip, "base64"))
        zip.extractAllTo(conf.apps_dir + '/' + name)
        res.ok({ name })
    }
    catch (ex) {
        res.error(ex.message)
    }
})

// TODO: Restrictions
router.delete('/:id', (req, res) => {
    try {
        rmSync(conf.apps_dir + '/' + req.params.id, { recursive: true, force: true })
        res.ok()
    }
    catch (ex) {
        res.error(ex.message)
    }
})

module.exports = router
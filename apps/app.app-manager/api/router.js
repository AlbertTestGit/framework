const express = require('express')
const router = express.Router()
const fs = require('fs')
const { resolve } = require('path')
const conf = require('./conf')

router.use('/:module/api/', (req, res, next) => {
    const filename = `${conf.apps_dir}/${req.params.module}/api/index.js`
    if (fs.existsSync(filename))
        require(filename)(req, res, next)
    else
        res.sendStatus(404)
})

// Hide system files
router.get('/index.js', (req, res) => res.sendStatus(404))
router.get('/*/_private*', (req, res) => res.sendStatus(404))

router.use((req, res, next) => {
    return express.static(conf.apps_dir + '/')(req, res, next)
})

module.exports = router
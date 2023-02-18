const express = require('express')
const router = express.Router()

const app_manager = 'app.app-manager'
const app_index = 'app.app-manager'

router.get('/', (req, res) => res.redirect(`./${app_index}/`))

router.use((req, res, next) => {
    return require(`./apps/${app_manager}/api/router.js`)(req, res, next)
})

module.exports = router
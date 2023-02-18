const express = require('express')
const app = express()

app.use(express.json())
app.use(require('./express-handler'))
app.use(require('./index-apps'))

app.listen(3000)
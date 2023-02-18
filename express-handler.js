const router = require('express').Router()

router.use('', (req, res, next) => {
	res.ok = data => res.send({ success: true, data })
	res.error = (data) => res.send({ success: false, data })
	res.handle = (async_func) => {
		async_func
			.then(r => res.ok(r))
			.catch(ex => {
				let msg = (ex.response && ex.response.data) || ex.message || ex
				console.log('Request failed', msg, ex.stack) // 
				res.error(msg)
			})
	}
	next()
})

module.exports = router


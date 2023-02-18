import { readFiles } from "./libs/file.js"
import './libs/mithril.min.js'

const DATA = {
    apps: []
}

function refreshList(){
    fetch("./api/").then(x => x.json()).then(rows => {
        DATA.apps = rows
        m.redraw()
    })
}

function openAddonInfo(name) {
    document.getElementById('addon_frame').src = `../${name}/info.html`
}

function dropPlugin(name) {
    return fetch(`./api/${name}`, { method: "DELETE" }).then(refreshList)
}

async function installPugins(input) {
    const files = await readFiles(input)
    const headers = { 'Content-Type': 'application/json' }
    fetch('./api', { headers, method: 'POST', body: JSON.stringify(files[0]) }).then(refreshList)
}

const app = {
    view(vnode) {
        const imp_label = m.trust('<b>Import plugin:</b>')
        const imp_panel = m('input', { type: 'file', accept: ".zip", multiple: 1 })
        const imp_button = m('button', { onclick() { installPugins(imp_panel.dom) } }, 'Import')
        return m('div', [imp_label, imp_panel, imp_button, DATA.apps.map(addonView)])
    }
}


function addonView(name) {
    // const open_link = m.trust(`<a href="./api/${name}/index.html">open</a>`)
    const open_link = m.trust(`<a href="../${name}/">open</a>`)
    const export_link = m.trust(`<a href="./api/${name}/@export">export</a>`)
    const info_link = m('a', { href: 'javascript:void(0)', onclick() { openAddonInfo(name) } }, 'info')
    const drop_link = m('a', { href: 'javascript:void(0)', onclick() { dropPlugin(name) } }, 'drop')

    return m('div', [name, open_link, info_link, export_link, drop_link])
}

m.mount(document.getElementById('apps'), app)

refreshList()
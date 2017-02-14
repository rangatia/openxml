const util = require('util')
const fs = require('fs')

const OpenXMLPackage = require('./openxmlpackage')

module.exports = class OpenXML {

    constructor () {
        this.Package = new OpenXMLPackage()
    }

    open (filename) {
        return new Promise( (resolve, reject) => {
            fs.readFile(filename, 'base64', (err, data) => {
                if (err) throw err
                this.Package.readPackage(data)
                    .then( pkg => {
                        this.Package = pkg
                        resolve()
                    })
            })
        })
    }

    save (filename) {
        this.Package.writePackage()
            .then( data => {
                fs.writeFile(filename, data, {encoding: 'base64'}, err => {
                    if (err) throw err
                })                    
            })
    }

}

// API Changes
// OpenXMLUtils not implemented
// OpenXML constants as OpenXMLX


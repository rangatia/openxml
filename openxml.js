const fs = require('fs')
const OpenXMLPackage = require('./openxmlpackage')

module.exports = class OpenXML {
  constructor () {
    this.Package = new OpenXMLPackage()
  }

  /**
   * read file contents to buffer
   */
  open (filename) {
    return new Promise((resolve, reject) => {
      fs.readFile(filename, (err, data) => {
        if (err) throw err
        this.Package.readPackage(data)
          .then(pkg => {
            this.Package = pkg
            resolve()
          })
      })
    })
  }

  /**
   * write zip object content (as buffer) to filesystem
   */
  save (filename) {
    this.Package.writePackage()
      .then(data => {
        fs.writeFile(filename, data, err => {
          if (err) throw err
        })
      })
  }

}

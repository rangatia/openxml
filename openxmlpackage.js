const JSZip = require('jszip')
const OpenXMLPart = require('./openxmlpart')
const OpenXMLRelationship = require('./openxmlrelationship')
const OpenXMLX = require('./openxmlx')

const Ltxml = require('./ltxml')
const DOMParser = require('xmldom').DOMParser
Ltxml.DOMParser = DOMParser
const XDocument = Ltxml.XDocument
const XElement = Ltxml.XElement
const XAttribute = Ltxml.XAttribute

module.exports = class OpenXMLPackage {

  constructor () {
    this.parts = {}
    this.ctXDoc = null
  }

  readPackage (data) {
    return new Promise((resolve, reject) => {
      const zip = new JSZip()
      zip.loadAsync(data)
        .then(zipData => {
          let constructParts = []
          Object.keys(zipData.files).map(f => {
            if (f.endsWith('/')) return
            const zipFile = zip.file(f)

            f = (f !== '[Content_Types].xml' ? '/' : '') + f
            const ext = f.substring(f.lastIndexOf('.') + 1)
            let asType = (ext === 'xml' || ext === 'rels') ? 'string' : 'base64'
            constructParts.push(
              zipFile.async(asType)
                .then(data => {
                  this.parts[f] = new OpenXMLPart(this, f, null, null, data)
                })
            )
          })
          Promise.all(constructParts)
            .then(() => {
              const ctf = this.parts['[Content_Types].xml']
              if (!ctf) throw new Error('Invalid OpenXML document. Content-Type item not found in package')

              this.ctXDoc = XDocument.parse(ctf.xml)

              for (const part in this.parts) {
                if (part === '[Content_Types].xml') continue

                const ct = this.getContentType(part)
                const thisPart = this.parts[part]
                thisPart.contentType = ct

                if (ct.endsWith('xml')) {
                  thisPart.xml = thisPart.data
                  thisPart.partType = 'xml'
                } else {
                  thisPart.base64 = thisPart.data
                  thisPart.partType = 'base64'
                }
              }

              resolve(this)
            })
        })
    })
  }

  writePackage () {
    return new Promise((resolve, reject) => {
      const zip = new JSZip()

      for (const part in this.parts) {
        let partContent = null

        const thisPart = this.parts[part]
        if (!thisPart) continue

        if (part === '[Content_Types].xml') {
          partContent = this.ctXDoc
          zip.file('[Content_Types].xml', partContent.toString(false))
          continue
        }

        let ct = null
        const cte = this.ctXDoc.getRoot()
          .elements(OpenXMLX.CT.Override)
          .firstOrDefault(e => e.attribute('PartName').value === part)
        if (!cte) {
          const ext = part.substring(part.lastIndexOf('.') + 1).toLowerCase()
          const dct = this.ctXDoc.getRoot()
            .elements(OpenXMLX.CT.Default)
            .firstOrDefault(e => e.attribute('Extension').value.toLowerCase() === ext)
          if (!dct) throw new Error('Invalid OpenXML document. Unable to process content type')
          ct = dct.attribute('ContentType').value
        } else ct = cte.attribute('ContentType').value

        let type = null
        let isLtxml = false

        if (ct.endsWith('xml')) {
          if (!thisPart.xDoc) partContent = thisPart.xml
          else {
            partContent = thisPart.xDoc
            isLtxml = true
          }
          type = 'xml'
        } else {
          partContent = thisPart.base64
          // type = 'binary'
          type = 'base64'
        }

        let name = part.charAt(0) === '/' ? part.substring(1) : part
        if (type !== 'xml') zip.file(name, partContent, { base64: true, compression: 'store' })
        else {
          if (isLtxml) zip.file(name, partContent.toString(false))
          else zip.file(name, partContent)
        }
      }
      zip.generateAsync({ type: 'nodebuffer', compression: 'deflate' })
        .then(content => resolve(content))
        .catch(err => reject(err))
    })
  }

  addPart (uri, contentType, partType, data) {
    const ctEl = this.ctXDoc.getRoot().elements(OpenXMLX.CT.Override)
      .firstOrDefault(or => or.attribute('PartName').value === uri)

    if (ctEl || this.parts[uri]) throw new Error('Invalid operation, trying to add a part that already exists')

    let newPart = new OpenXMLPart(this, uri, contentType, partType, data)
    this.parts[uri] = newPart

    if (contentType === OpenXMLX.contentTypes.relationships) return newPart
    this.ctXDoc.getRoot().add(
      new XElement(OpenXMLX.CT.Override,
        new XAttribute('PartName', uri),
        new XAttribute('ContentType', contentType)))

    return newPart
  }

  deletePart (part) {
    const uri = part.uri
    this.parts[uri] = null

    const ctEl = this.ctXDoc.getRoot().elements(OpenXMLX.CT.Override)
      .firstOrDefault(or => or.attribute('PartName').value === uri)
    if (ctEl) ctEl.remove()
  }

  addRelationship (relationshipId, relationshipType, target, targetMode) {
    if (!targetMode) targetMode = 'Internal'
    let rootRelationshipPart = this.getPartByUri('/_rels/.rels')
    if (!rootRelationshipPart) {
      rootRelationshipPart = this.addPart('/_rels/.rels', OpenXMLX.contentTypes.relationships, 'xml',
        new XDocument(
          new XElement(OpenXMLX.PKGREL.Relationships,
            new XAttribute('xmlns', OpenXMLX.pkgRelNs.namespaceName))))
    }
    OpenXMLPart.addRelationshipToRelPart(rootRelationshipPart, relationshipId, relationshipType, target, targetMode)
  }

  deleteRelationship (relationshipId) {
    const rootRelationshipsPart = this.getPartByUri('/_rels/.rels')
    const rxDoc = rootRelationshipsPart.getXDocument()
    const rxe = rxDoc.getRoot().elements(OpenXMLX.PKGREL.Relationship)
      .firstOrDefault(r => r.attribute('Id').value === relationshipId)
    if (rxe) rxe.remove()
  }

  getContentType (uri) {
    const ct = this.ctXDoc
      .descendants(OpenXMLX.CT.Override)
      .firstOrDefault(o => o.attribute('PartName').value === uri)

    if (!ct) {
      const ext = uri.substring(uri.lastIndexOf('.') + 1)
      const dct = this.ctXDoc
        .descendants(OpenXMLX.CT.Default)
        .firstOrDefault(d => d.attribute('Extension').value === ext)
      if (dct) return dct.attribute('ContentType').value
      else return null
    }
    return ct.attribute('ContentType').value
  }

  getPartById (relationshipId) {
    const rel = this.getRelationshipById(relationshipId)
    if (rel) return this.getPartByUri(rel.targetFullName)
    return null
  }

  getPartByRelationshipType (relationshipType) {
    const parts = this.getPartsByRelationshipType(relationshipType)
    if (parts.length < 1) return null
    return parts[0]
  }

  getPartByUri (uri) {
    return this.parts[uri]
  }

  getParts () {
    const parts = []
    const rels = this.getRelationships()
    for (const rel of rels) {
      if (rel.targetMode === 'Internal') parts.push(this.getPartByUri(rel.targetFullName))
    }
    return parts
  }

  getPartsByRelationshipType (relationshipType) {
    const parts = []
    const rels = this.getRelationshipsByRelationshipType(relationshipType)
    for (const rel of rels) {
      parts.push(this.getPartByUri(rel.targetFullName))
    }
    return parts
  }

  getPartsByContentType (contentType) {
    const parts = []
    const rels = this.getRelationshipsByContentType(contentType)
    for (const rel in rels) {
      parts.push(this.getPartByUri(rel.targetFullName))
    }
    return parts
  }

  getRelationshipById (relationshipId) {
    const rels = this.getRelationships()
    for (const rel of rels) {
      if (rel.relationshipId === relationshipId) return rel
    }
    return null
  }

  getRelationships () {
    const rootRelationshipsPart = this.getPartByUri('/_rels/.rels')
    return this.getRelationshipsFromPart(null, rootRelationshipsPart)
  }

  getRelationshipsFromPart (part, relsPart) {
    const pkg = part ? null : this
    const rxDoc = relsPart.getXDocument()
    const rels = rxDoc.getRoot().elements(OpenXMLX.PKGREL.Relationship)
      .select(r => {
        const targetMode = r.attribute('TargetMode') ? r.attribute('TargetMode').value : null
        return new OpenXMLRelationship(
          pkg,
          part,
          r.attribute('Id').value,
          r.attribute('Type').value,
          r.attribute('Target').value,
          targetMode
        )
      })
      .toArray()
    return rels
  }

  getRelationshipsByRelationshipType (relationshipType) {
    const rootRelationshipsPart = this.getPartByUri('/_rels/.rels')
    const rxDoc = rootRelationshipsPart.getXDocument()
    // var that = this;
    const rels = rxDoc.getRoot().elements(OpenXMLX.PKGREL.Relationship)
      .where(r => r.attribute('Type').value === relationshipType)
      .select(r => {
        const targetMode = r.attribute('TargetMode') ? r.attribute('TargetMode').value : null
        return new OpenXMLRelationship(
          this, // that,
          null,
          r.attribute('Id').value,
          relationshipType,
          r.attribute('Target').value,
          targetMode
        )
      })
      .toArray()
    return rels
  }

  getRelationshipsByContentType (contentType) {
    const rootRelationshipsPart = this.getPartByUri('/_rels/.rels')
    const rels = []
    if (rootRelationshipsPart) {
      // const allRels = getRelationshipsFromRelsXml(this, null, rootRelationshipsPart)
      const allRels = this.getRelationshipsFromPart(null, rootRelationshipsPart)
      for (const rel of allRels) {
        if (rel.targetMode === 'External') continue
        let ct = this.getContentType(rel.targetFullName)
        if (ct !== contentType) continue
        rels.push(rel)
      }
    }
    return rels
  }

}

/*
 changelog
 |==============================|=========================================
 | Legacy Code                  | Changes
 |==============================|=========================================
 | openFromBase64Internal       | readPackage
 | openFromFlatOpcInternal      | not handling flatOPC
 | ()                           | constructor
 | openFromBase64               | directly piped to writePackage
 | openFromFlatOpc              | not handling flatOPC
 | saveToBase64                 | writePackage
 | saveToFlatOpc                | not handling flatOPC
 | addPart                      |
 | deletePart                   |
 | function getRelationshipsFromRelsXml | getRelationshipsFromPart
 | getRelationships             |
 | getParts                     |
 | getRelationshipsByRelationshipType |
 | getPartsByRelationshipType   |
 | getPartByRelationshipType    |
 | getRelationshipsByContentType|
 | getPartsByContentType        |
 | getRelationshipById          |
 | getPartById                  |
 | getPartByUri                 |
 | function addRelationshipToRelPart    | OpenXMLPart.addRelationshipToRelPart
 | addRelationship              |
 | deleteRelationship           |
 | getContentType               |
 | *Part(s)                     | userland functions
 |==============================|=========================================
 */

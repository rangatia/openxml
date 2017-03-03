const OpenXMLX = require('./openxmlx')
// const OpenXMLRelationship = require('./openxmlrelationship')
const Ltxml = require('./ltxml')
const XDocument = Ltxml.XDocument
const XElement = Ltxml.XElement
const XAttribute = Ltxml.XAttribute

module.exports = class OpenXMLPart {
  constructor (pkg, uri, contentType, partType, content) {
    this.pkg = pkg
    this.uri = uri
    this.contentType = contentType
    this.partType = partType
    this.data = null
    this.buffer = null
    this.xml = null
    this.xDoc = null

    partType = (uri === '[Content_Types].xml' ? 'xml' : partType)

    if (partType) {
      if (partType === 'buffer') this.buffer = content
      else if (partType === 'xml') {
        if (content.nodeType === 'Element' || content.nodeType === 'Document') this.xDoc = content
        else this.xml = content
      }
    } else this.data = content
  }

  getXDocument () {
    if (!this.xDoc) this.xDoc = XDocument.parse(this.xml)
    return this.xDoc
  }

  getParts () {
    let parts = []
    const rels = this.getRelationships()
    for (const rel of rels) {
      if (rel.targetMode === 'Internal') parts.push(this.pkg.getPartByUri(rel.targetFullName))
    }
    return parts
  }

  getPartsByRelationshipType (relationshipType) {
    let parts = []
    const rels = this.getRelationships().filter(r => r.relationshipType === relationshipType)
    for (const rel of rels) {
      parts.push(this.pkg.getPartByUri(rel.targetFullName))
    }
    return parts
  }

  getPartsByContentType (contentType) {
    let parts = []
    const rels = this.getRelationshipsByContentType(contentType)
    for (const rel of rels) {
      parts.push(this.pkg.getPartByUri(rel.targetFullName))
    }
    return parts
  }

  addRelationshipToPart (relId, relType, target, targetMode) {
    const rxDoc = this.getXDocument()
    const tm = (targetMode !== 'Internal' ? new XAttribute('TargetMode', 'External') : null)
    rxDoc.getRoot().add(
      new XElement(OpenXMLX.PKGREL.Relationship,
        new XAttribute('Id', relId),
        new XAttribute('Type', relType),
        new XAttribute('Target', target),
        tm))
  }

  getRelationships () {
    let rels = []
    const relsPartUri = this.uri.substring(0, this.uri.lastIndexOf('/')) +
      '/_rels/' + this.uri.substring(this.uri.lastIndexOf('/') + 1) + '.rels'
    const relsPart = this.pkg.getPartByUri(relsPartUri)
    if (relsPart) rels = this.pkg.getRelationshipsFromPart(this, relsPart)
    return rels
  }

  getRelationshipsByContentType (contentType) {
    let rels = []
    const allRels = this.getRelationships()
    for (const rel of allRels) {
      if (rel.targetMode !== 'Internal') continue
      const ct = this.pkg.getContentType(rel.targetFullName)
      if (ct === contentType) rels.push(rel)
    }
    return rels
  }

  addRelationship (relId, relType, target, targetMode) {
    const relsPartUri = this.uri.substring(0, this.uri.lastIndexOf('/')) +
      '/_rels/' + this.uri.substring(this.uri.lastIndexOf('/') + 1) + '.rels'
    let relsPart = this.pkg.getPartByUri(relsPartUri)
    if (!relsPart) {
      relsPart = this.pkg.addPart(relsPartUri, OpenXMLX.contentTypes.relationships, 'xml',
        new XDocument(
          new XElement(OpenXMLX.PKGREL.Relationships,
            new XAttribute('xmlns', OpenXMLX.pkgRelNs.namespaceName))))
    }
    this.addRelationshipToPart(relId, relType, target, targetMode)
  }

  deleteRelationship (relId) {
    const relsPartUri = this.uri.substring(0, this.uri.lastIndexOf('/')) +
      '/_rels/' + this.uri.substring(this.uri.lastIndexOf('/') + 1) + '.rels'
    const relsPart = this.pkg.getPartByUri(relsPartUri)
    if (relsPart) {
      const relsPartXDoc = relsPart.getXDocument()
      const theRel = relsPartXDoc.getRoot().elements(OpenXMLX.PKGREL.Relationships)
        .firstOrDefault(r => r.attribute('Id').value === relId)

      if (theRel) theRel.remove()
    }
  }

}

/*
 changelog
 |==============================|=========================================
 | Legacy Code                  | Changes
 |==============================|=========================================
 | ()                           | constructor
 | getXDocument                 |
 | putXDocument                 | not used
 | getRelsPartUriOfPart         | inline
 | getRelsPartOfPart            | inline
 | getRelationships             |
 | getParts                     |
 | getRelationshipsByRelationshipType | getRelations().filter( r.relationType )
 | getPartsByRelationshipType   |
 | getPartByRelationshipType    | as above. array index 0
 | getRelationshipsByContentType|
 | getPartsByContentType        |
 | getRelationshipById          | getRleations().filter( r.relationshipId)
 | getPartById                  | getRelations.filter(by rId).getPartByURI(r.targetFullName)]
 | addRelationship              |
 | deleteRelationship           |
 | *part                        | userland functions
 | OpenXMLPackage.addRelationshipToRelPart  | addRelationshipToPart
 */

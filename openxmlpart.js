const Ltxml = require('./ltxml')
const XDocument = Ltxml.XDocument
const XElement = Ltxml.XElement
const OpenXMLRelationship = require('./openxmlrelationship')
const openXMLX = require('./openxmlx')
const XAttribute = Ltxml.XAttribute

module.exports = class OpenXMLPart {

    constructor (pkg, uri, contentType, partType, data) {
        // console.log(`uri ${uri}`)
        this.pkg = pkg
        this.uri = uri
        this.contentType = contentType
        this.partType = partType
        this.data = null
        this.base64 = null
        this.xml = null
        this.xDoc = null

        partType = (uri === "[Content_Types].xml" ? 'xml' : partType)

        if (partType) {
            if (partType === 'base64') this.base64 = data
            else if (partType === 'xml') {
                if (data.nodeType === 'Element' || data.nodeType === 'Document') this.xDoc = data
                else this.xml = data
            }
        } else this.data = data
    }

    getXDocument () {
        if (!this.xDoc) this.xDoc = XDocument.parse(this.xml) // XDocument.parse(decodeURIComponent(this.xml))
        return this.xDoc
    }

    getParts () {
        let parts = []
        const rels = this.getRelationships()
        for (const rel of rels) {
            if (rel.targetMode === "Internal") parts.push(this.pkg.getPartByUri(rel.targetFullName))
        }
        return parts
    }

    getPartsByRelationshipType (relationshipType) {
        let parts = []
        // const rels = this.getRelationshipsByRelationshipType(relationshipType)
        const rels = this.getRelationships().filter( r => r.relationshipType === relationshipType)
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

    addRelationshipToPart(relId, relType, target, targetMode) {
        const rxDoc = this.getXDocument()
        const tm = (targetMode !== "Internal" ? new XAttribute("TargetMode", "External") : null )
        rxDoc.getRoot().add(
            new XElement(openXMLX.PKGREL.Relationship,
                new XAttribute("Id", relationshipId),
                new XAttribute("Type", relationshipType),
                new XAttribute("Target", target),
                tm))
    }

    getRelationships () {
        let rels = []
        const relsPartUri = this.uri.substring(0, this.uri.lastIndexOf('/'))
            + '/_rels/' + this.uri.substring(this.uri.lastIndexOf('/') + 1)
            + '.rels'
        const relsPart = this.pkg.getPartByUri(relsPartUri)
        if (relsPart) rels = this.pkg.getRelationshipsFromPart(this, relsPart)
        return rels
    }

    getRelationshipsByContentType (contentType) {
        let rels = []
        const allRels = this.getRelationships()
        for (rel of allRels) {
            if (rel.targetMode !== 'Internal') continue
            const ct = this.pkg.getContentType(rel.targetFullName)
            if (ct === contentType) rels.push(rel)
        }
        return rels
    }

    addRelationship (relId, relType, target, targetMode) {
        const relsPartUri = this.uri.substring(0, this.uri.lastIndexOf('/'))
            + '/_rels/' + this.uri.substring(this.uri.lastIndexOf('/') + 1)
            + '.rels'
        const relsPart = this.pkg.getPartByUri(relsPartUri)
        if (!relsPart) {
            relsPart = this.pkg.addPart(relsPartUri, openXMLX.contentTypes.relationships, 'xml',
                new XDocument(
                    new XElement(openXMLX.PKGREL.Relationships,
                        new XAttribute('xmlns', openXMLX.pkgRelNs.namespaceName))))
        }
        this.addRelationshipToPart(relId, relType, target, targetMode)
    }

    deleteRelationship (relId) {
        const relsPartUri = this.uri.substring(0, this.uri.lastIndexOf('/'))
            + '/_rels/' + this.uri.substring(this.uri.lastIndexOf('/') + 1)
            + '.rels'
        const relsPart = this.pkg.getPartByUri(relsPartUri)
        if (relsPart) {
            const relsPartXDoc = relsPart.getXDocument()
            const theRel = relsPartXDoc.getRoot().elements(openXMLX.PKGREL.Relationships)
                .firstOrDefault( r => r.attribute('Id').value === relId )

            if (theRel) theRel.remove()
        }
    }

}

/**
 * Changes
 * 
 * 
 */
// API Changes [I=implemented | NR<reason>=not required | T<new method>=Transitioned to]
// [T] ()                               // constructor
// [I] getXDocument
// [N] putXDocument                     // not used anywhere
// [T] getRelsPartUriOfPart             // inline code
// [T] getRelsPartOfPart                // inline code
// [I] getRelationships
// [I] getParts
// [N] getRelationshipsByRelationshipType   // getRelations().filter( r.relationType )
// [I] getPartsByRelationshipType
// [N] getPartByRelationshipType            // use above method with array index 0
// [I] getRelationshipsByContentType        
// [I] getPartsByContentType
// [N] getRelationshipById                  // getRleations().filter( r.relationshipId)
// [N] getPartById                          // getRelations.filter(by rId).getPartByURI(r.targetFullName)]
// [I] addRelationship
// [I] deleteRelationship
// [N] *part                                // userland functions
// 
// from OpenXMLPackage
// [T] addRelationshipToRelPart             // as this.addRelationshipToPart
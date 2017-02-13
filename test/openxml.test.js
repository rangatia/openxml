'use strict'

const expect = require('chai').expect

const OpenXML = require('../openxml')
const OpenXMLPackage = require('../openxmlpackage')

describe('class OpenXML', () => {
    describe('new OpenXML()', () => {
        const o = new OpenXML()
        it('should create new OpenXML with its Package property as new OpenXMLPackage', () => {
            expect(o).to.be.instanceof(OpenXML)
                .that.to.have.property('Package')
            expect(o.Package).to.be.instanceof(OpenXMLPackage)
        });
    });

    describe('open', () => {
        // should throw error if no filename
        // should only process OpenXML formats
        // should create XDocument structure from file        
    });

    describe('save', () => {
        // throw error if no filename
        // throw error if not OpenXML document format
        // write to fs        
    });
});
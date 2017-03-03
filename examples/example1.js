/**
 * Basic use-case.
 * 
 * Open & save a blank pptx file.
 * 

const OpenXML = require('../openxml')

const pptx = new OpenXML()
pptx.open('./examples/blank.pptx')
    .then( () => pptx.save('./examples/out.pptx'))

 */
// const OpenXML = require('../openxml')
// // const o = new OpenXML()
// // o.loadAsync('./examples/blank.pptx')
// //     .then(p => console.log(p))

// const pptx = new OpenXML().loadAsync('./examples/blank.pptx')
// // pptx.then(p => console.log(p))

/* 2 */
const fs = require('fs')
const OpenXML = require('../openxml')
// console.log(OpenXML.CT.Override)
const pptx = new OpenXML.OpenXMLPackage()
pptx.loadAsync('./examples/blank.pptx')
    .then(() => {
        pptx.generateAsync('./examples/out.pptx')
    })


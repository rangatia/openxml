/**
 * Basic use-case.
 * 
 * Open & save a blank pptx file.
 * 
 */

const OpenXML = require('../openxml')

const pptx = new OpenXML()
pptx.open('./examples/blank.pptx')
    .then( () => pptx.save('./examples/out.pptx'))


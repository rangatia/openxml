const OpenXML = require('../openxml')
const pptx = new OpenXML.OpenXMLPackage()
pptx.readPackage('./examples/blank.pptx')
    .then(() => {
        pptx.writePackage('./examples/out.pptx')
    })


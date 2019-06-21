'use strict';

// get the file to open, which is the 3rd argument on the command line
const fileToOpen = process.argv[2];
if (fileToOpen === undefined) { // quit immediately if there is no 3rd argument
    console.log('Usage: node parse.js filename');
    console.log('Error: must provide filename to parse');
    return;
}

// put document.xml into doc as a string
const AdmZip = require('adm-zip');
let zip = new AdmZip('docs/' + fileToOpen);
let doc;
zip.getEntries().forEach(zipEntry => {
    if (zipEntry.entryName == 'word/document.xml')
        doc = zipEntry.getData().toString('utf8');
});

const convert = require('xml-js');
let i = 0;
let options = {
    compact: true,
    spaces: 4,
    ignoreDeclaration: true,
    ignoreInstruction: true,
    ignoreComment: true,
    ignoreCdata: true,
    ignoreDoctype: true,
    /* attributesKey: 'type', textKey: 'value', */
    /* attributeValueFn: function (value, name) {
        if (name != 'w:t') return '';
        if (value != 'w:val') {
            // console.log(i++ + ' ' + value);
            return 'wat';
        }
    }, */
    /* attributeNameFn: function(name, value) {
        if (name == 'w:p') {
            console.log(i++ + ' ' + name);
            return 'wot';
        }
    } */
};
let docObj = convert.xml2js(doc, options); // convert xml to js object

// remove non-text xml tags from the converted js object
function removeTags(docObj) {
    if ('w:document' in docObj) {
        if ('_attributes' in docObj['w:document']) {
            // remove namespaces
            delete docObj['w:document']['_attributes'];
        }
        if ('w:body' in docObj['w:document']) {
            if ('w:sectPr' in docObj['w:document']['w:body']) {
                delete docObj['w:document']['w:body']['w:sectPr'];
            }
        }
    }
}

removeTags(docObj);

const fs = require('fs');
fs.writeFile('test.json', JSON.stringify(docObj, null, 4), (err) => { if (err) throw err; });

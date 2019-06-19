'use strict';

// get the file to open, which is the 3rd argument on the command line
const fileToOpen = process.argv[2];
if (fileToOpen === undefined) { // quit immediately if there is no 3rd argument
    console.log('Usage: node parse.js filename');
    console.log('Error: must provide filename to parse');
    return;
}

// extract the specified file to a directory '/file' in the current working directory
const extract = require('extract-zip');
extract('docs/' + fileToOpen, { dir: process.cwd() + '/file' }, function (err) {
    if (err) throw err;
});

const convert = require('xml-js');
const fs = require('fs');

let xml = fs.readFileSync('file/word/document.xml', 'utf-8');
let options = { compact: true, spaces: 4, ignoreDeclaration: true, ignoreInstruction: true, ignoreComment: true, ignoreCdata: true, ignoreDoctype: true, attributesKey: 'type', textKey: 'value' };
let result = convert.xml2json(xml, options);
// console.log(result);

fs.writeFile('test.json', result, () => { });

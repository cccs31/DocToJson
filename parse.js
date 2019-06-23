'use strict';

// get the file to open, which is the 3rd argument on the command line
const fileToOpen = process.argv[2];
if (fileToOpen === undefined) {
    console.log('Usage: node parse.js filename');
    console.log('Error: must provide filename to parse');
    return;
}
if (!fileToOpen.endsWith('.docx')) {
    console.log('Error: filetype must be .docx');
    return;
}

// put document.xml into doc as a string
const AdmZip = require('adm-zip');
let zip = new AdmZip('docs/' + fileToOpen);
let doc; // a string containing document.xml
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
};
let docObj = convert.xml2js(doc, options); // convert xml to js object

// final javascript object with only the necessary properties
let result = {
    filename: fileToOpen,
    content: [] // each index stores one text block
};

// An object that holds a name, value, and optional bible verse.
//
// Each instance of ContentObj stores one text block, which represents all the
// text and text properties of a single W_R object.
// name and value represents the text type and the text string, respectively.
// bible is only set if value contains bible verse references, and is left
// undefined otherwise.
function ContentObj(name, value) {
    this.name = name;
    this.value = value;
    this.bible;
};

findTextBlocks(docObj);

// Finds all the W_R objects in docObj and calls copyTextAndProperties(W_R) to
// copy the appropriate text and text properties to create a ContentObj object.
// A ContentObj object is returned, which is stored as an element in
// result.content.
function findTextBlocks(docObj) {
    const W_DOCUMENT = 'w:document';
    const W_BODY = 'w:body';
    const W_P = 'w:p';
    const W_R = 'w:r'; // a text block; one ContentObj per W_R

    if (W_DOCUMENT in docObj && W_BODY in docObj[W_DOCUMENT]
        && W_P in docObj[W_DOCUMENT][W_BODY]) {

        let index = 0;
        // iterate over W_P
        for (let i in docObj[W_DOCUMENT][W_BODY][W_P]) {
            // Look for W_R, which contains the text and their text properties,
            // at each index of W_P.
            if (W_R in docObj[W_DOCUMENT][W_BODY][W_P][i]) {
                console.log(index++);

                let textBlock
                    = copyTextAndProperties(
                        docObj[W_DOCUMENT][W_BODY][W_P][i][W_R]);

                if (textBlock != null) {
                    // insert textBlock   
                    // result.content.splice(index, 0, textBlock);
                }
            }
        }
    }
}

// Finds and copies the text properties and text in W_R to two arrays, names
// and values, which are then used to create a ContentObj object.
//Creates one ContentObj object, which holds the name, value, and
// (computed) bible verse from each wR object.
//
// W_R represents one text block, which may contain several different text
// types. If W_R contains more than one text type, the value of the created
// ContentObj object will be an array that contains more ContentObj objects.
//
// If the previous element in W_R has the same text properties as the current
// element, the text of the current element is concatenated with the text of
// the previous element.
function copyTextAndProperties(W_R) {
    const W_T = 'w:t'; // the text object
    const W_RPR = 'w:rPr'; // contains all the text properties
    const W_B = 'w:b'; // text property - bold
    const W_I = 'w:i'; // text property - italics
    const W_SZ = 'w:sz'; // text property - font size
    const W_U = 'w:u'; // text property - underline
    const ATTR = '_attributes';
    const TEXT = '_text'; // text string

    let names = [], values = []; // parameters for ContentObj
    let W_RIndex = 0;
    let sizeW_R = Object.keys(W_R).length;

    // iterate over W_R
    for (let i in W_R) {
        // find the text properties in W_RPR and the text in W_T
        if (W_RPR in W_R[i] && W_T in W_R[i]) {
            console.log('W_RIndex ' + W_RIndex);
            // console.log(W_R[i][W_RPR]);

            let isDone = false;

            // Iterate over wRpr, the text properties, and create a ContentObj
            // object based on the text properties.
            for (let j in W_R[i][W_RPR]) {
                switch (j) {
                    case W_SZ:
                        let W_VAL = 'w:val';
                        if (ATTR in W_R[i][W_RPR][W_SZ] && W_VAL in W_R[i][W_RPR][W_SZ][ATTR]) {
                            const MIN_TITLE_SIZE = 28;
                            if (W_R[i][W_RPR][W_SZ][ATTR][W_VAL] >= MIN_TITLE_SIZE) {
                                names[W_RIndex] = 'title';
                                isDone = true;
                            }
                        }
                        break;
                    case W_B:
                        if (names[W_RIndex] === undefined)
                            names[W_RIndex] = '';
                        names[W_RIndex] += 'bold';
                        break;
                    case W_I:
                        if (names[W_RIndex] != undefined)
                            names[W_RIndex] += '_italic';
                        if (names[W_RIndex] === undefined)
                            names[W_RIndex] = 'italic';
                        break;
                    case W_U:
                        if (names[W_RIndex] != undefined)
                            names[W_RIndex] += '_underline';
                        if (names[W_RIndex] === undefined)
                            names[W_RIndex] = 'underline';
                        break;
                    default:
                        break;
                }
                // text of type 'title' should not have other properties
                if (isDone) break;
            }
            if (!isDone)
                (names[W_RIndex] != undefined) ? names[W_RIndex] += '_text' : names[W_RIndex] = 'text';

            // get the text string in wT and put it in the ContentObj object
            if (TEXT in W_R[i][W_T]) {
                // console.log(W_R[i][W_T][TEXT]);
                values[W_RIndex] = W_R[i][W_T][TEXT];
            } else { // skip the current index of W_R if there is no text
                console.log('---SKIPPED DUE TO NO TEXT---');
                names.pop();
                continue;
            }

            if (concatSameProperties(names, values, W_RIndex))
                --W_RIndex;
            console.log(names, values); // one iteration of the array of wR

            // if (name == textBlock.name) {
            //     // concatenate text with the same properties
            //     textBlock.value += value;
            // } else { // text has different properties
            //     if (sizeW_R > 1) {
            //         textBlock.name = 'paragraph';
            //         textBlock.value.push(value);
            //     }
            //     // one text type in the entire text block
            //     textBlock.name = name;
            //     textBlock.value = value;
            // }

            // display block on every iteration c of W_R
            // console.log('block:');
            // console.log(textBlock);
            // iterate(textBlock);
        }
        W_RIndex++;
    }

    console.log('---BLOCK---BLOCK---BLOCK---');
    console.log();
    // create textBlock, pass parameters names and values
    return createContentObj(names, values);
}

// Concatenate the element at values[index] with the previous element
// values[index] only if they have the same name (ie, same text properties).
// Returns true if concatenation was performed, false otherwise.
function concatSameProperties(names, values, index) {
    for (let mergeIndex = index; mergeIndex >= 0; --mergeIndex) {
        if (names[mergeIndex] == names[mergeIndex - 1]) {
            values[mergeIndex - 1] += values[mergeIndex];
            names.pop();
            values.pop();
            return true;
        } else {
            return false;
        }
    }
}

// Creates a ContentObj object from the values of the specified arrays names
// and values, which contain text properties and text strings from a W_R array.
//
// If bible verse references exist in the text strings in values, they are
// parsed and additionally stored in ContentObj.verses.
// Returns the newly created ContentObj object.
function createContentObj(names, values) {

}

const fs = require('fs');
fs.writeFile(
    'test.json',
    JSON.stringify(result, (key, value) => { if (value != null) return value; }, 4),
    (err) => { if (err) throw err; }
);

// iterate over all properties of the specified object obj
function iterate(obj) {
    for (let property in obj) {
        if (obj.hasOwnProperty(property)) {
            if (typeof obj[property] == 'object')
                iterate(obj[property]);
            else
                console.log(property + ': ' + obj[property]);
        }
    }
}

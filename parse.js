'use strict';

// get the file to open, which is the 3rd argument on the command line
const FILE_TO_OPEN = process.argv[2];
if (FILE_TO_OPEN === undefined) {
    console.log('Usage: node parse.js filename');
    console.log('Error: must provide filename to parse');
    return;
}
if (!FILE_TO_OPEN.endsWith('.docx')) {
    console.log('Error: filetype must be .docx');
    return;
}

// put document.xml into doc as a string
const AdmZip = require('adm-zip');
let zip = new AdmZip('docs/' + FILE_TO_OPEN);
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
    filename: FILE_TO_OPEN,
    content: [] // each index stores one text block
};

findAndInsertTextBlocks(docObj);

// write to JSON file
const fs = require('fs');
fs.writeFile(
    FILE_TO_OPEN.replace('.docx', '') + '.json',
    JSON.stringify(result, null, 4),
    (err) => { if (err) throw err; }
);

///////////////////////////// FUNCTION DEFINITIONS ////////////////////////////

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

function BibleVerse(verse, book) {
    this.book = book;
    this.verse = verse;
}

// Finds all the W_R objects in docObj and calls copyTextAndProperties(W_R) to
// copy the appropriate text and text properties to create a ContentObj object,
// which is inserted into result.content.
function findAndInsertTextBlocks(docObj) {
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
                let textBlock = copyTextAndProperties(
                    docObj[W_DOCUMENT][W_BODY][W_P][i][W_R]);

                if (textBlock != null) {
                    result.content.push(textBlock);
                }
            }
        }
    }
}

// Finds and copies the text properties and text in W_R to two arrays, names
// and values, which are then used to create a ContentObj object.
//
// W_R represents one text block, which may contain several different text
// types. If W_R contains more than one text type, the value of the created
// ContentObj object will be an array that contains other ContentObj objects.
//
// If the previous element in W_R has the same text properties as the current
// element, the text of the current element is concatenated with the text of
// the previous element.
//
// Returns a ContentObj object containing the text properties and text in W_R.
function copyTextAndProperties(W_R) {
    const W_T = 'w:t'; // text object
    const W_RPR = 'w:rPr'; // contains all the text properties
    const W_B = 'w:b'; // text property - bold
    const W_I = 'w:i'; // text property - italics
    const W_SZ = 'w:sz'; // text property - font size
    const W_U = 'w:u'; // text property - underline
    const ATTR = '_attributes';
    const TEXT = '_text'; // text string

    let names = [], values = []; // parameters for ContentObj
    let index = 0; // index of names and values

    // Text blocks that have only one text type in W_R are not arrays, and
    // should not be iterated over like arrays.
    let zeroLength = false;
    if (W_R.length == undefined) zeroLength = true;

    // iterate over W_R
    for (let i in W_R) {
        // find only the text properties in W_RPR and the text in W_T
        if ((W_RPR in W_R[i] && W_T in W_R[i])
            || (W_RPR in W_R && W_T in W_R)) {

            let isDone = false;

            // Iterate over W_RPR, the text properties, and put the necessary
            // text properties in name.
            for (let j in W_R[i][W_RPR]) {
                switch (j) {
                    case W_SZ:
                        let W_VAL = 'w:val';
                        if (ATTR in W_R[i][W_RPR][W_SZ]
                            && W_VAL in W_R[i][W_RPR][W_SZ][ATTR]) {

                            const MIN_TITLE_SIZE = 28;
                            if (W_R[i][W_RPR][W_SZ][ATTR][W_VAL]
                                >= MIN_TITLE_SIZE) {

                                names[index] = 'title';
                                isDone = true;
                            }
                        }
                        break;
                    case W_B:
                        if (names[index] != undefined)
                            names[index] += '_bold';
                        else
                            names[index] = 'bold';
                        break;
                    case W_I:
                        if (names[index] != undefined)
                            names[index] += '_italic';
                        else
                            names[index] = 'italic';
                        break;
                    case W_U:
                        if (names[index] != undefined)
                            names[index] += '_underline';
                        else
                            names[index] = 'underline';
                        break;
                    default:
                        break;
                }
                // text of type 'title' should not have other properties
                if (isDone) break;
            }
            if (!isDone) {
                (names[index] != undefined)
                    ? names[index] += '_text'
                    : names[index] = 'text';
            }

            // get the text string in wT and put it in values
            if (zeroLength && TEXT in W_R[W_T]) {
                values[index] = W_R[W_T][TEXT];
            } else if (!zeroLength && TEXT in W_R[i][W_T]) {
                values[index] = W_R[i][W_T][TEXT];
            } else { // skip the current index of W_R if there is no text
                names.pop();
                if (zeroLength)
                    break;
                continue;
            }

            if (concatenateSameProperties(names, values, index))
                --index;

            index++;
        }
        if (zeroLength) break; // iterate only once if W_R contains one object
    }

    return createContentObj(names, values);
}

// Concatenate the current element with the previous element only if they have
// the same name (ie, same text properties).
// Returns true if concatenation was performed, false otherwise.
function concatenateSameProperties(names, values, index) {
    if (names[index] == names[index - 1]) {
        values[index - 1] += values[index];
        names.pop();
        values.pop();
        return true;
    } else {
        return false;
    }
}

// Creates a ContentObj object from the values of the specified arrays names
// and values, which contain text properties and text strings from a W_R array.
//
// If bible verse references exist in the text strings in values, they are
// parsed and additionally stored in ContentObj.verses.
// Returns the newly created ContentObj object.
function createContentObj(names, values) {
    if (names.length == 0 || values.length == 0) return null;
    let textBlock = new ContentObj();

    // text is title
    if (names.length == 1 && names[0] == 'title') {
        // console.log('---TITLE---');
        textBlock.name = names[0];
        textBlock.value = values[0];
    }

    // text is a paragraph
    if (names[0] != 'title') {
        textBlock.name = 'paragraph';
        textBlock.value = [];
        for (let i in names) {
            let nestedBlock = new ContentObj(names[i], values[i]);
            nestedBlock = parseVerses(nestedBlock);
            textBlock.value.push(nestedBlock);
        }
    }

    return textBlock;
}

// Parses the specified ContentObj object textBlock to extract bible verse
// references into textBlock.bible.
// If there are no verses to extract, textBlock.bible is left undefined.
// Returns the specified object textBlock.
function parseVerses(textBlock) {
    let verses = [];
    let parenBlocks = textBlock.value.match(/\([^\)]+\)/g); // find verses

    if (parenBlocks != null) {
        for (let block of parenBlocks) {
            // not a bible verse if the parentheses block does not have a colon
            if (!block.includes(':'))
                continue;

            block = block.replace(/\(|\)/g, ''); // remove parentheses
            let singleVerses = block.split(/,|;/g); // remove delimiters
            for (let verse of singleVerses) {
                verse = verse.trim(); // remove whitespace
                verses.push(new BibleVerse(...verse.split(' ').reverse()));
            }
        }
    }

    textBlock.bible = (verses.length == 0) ? undefined : verses;
    return textBlock;
}

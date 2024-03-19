"use strict";
/* import dependencies - not sure any of this is necessary
/// <reference types="office-js" />
import * as Office from 'office.js';
*/
// global declarations
const html = document;
const runButton = html.getElementById("runButton");
const textField = html.getElementById("textField");
// check that office is ready
Office.onReady(() => {
    // check that the html has loaded
    html.addEventListener("DOMContentLoaded", () => {
        runButton.addEventListener("click", () => {
            Word.run(main) // run main function
                .catch((log) => console.error(log)); // unless it breaks
        });
    });
});
// run button function goes here
async function main(context) {
    const content = context.document.body;
    if (textField.value === "") {
        content.insertParagraph("Who goes there?", "End");
    }
    else {
        content.insertParagraph("Hello " + textField.value, "End");
    }
    await context.sync();
}

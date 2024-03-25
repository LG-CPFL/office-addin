"use strict";
// global declarations
const html = document;
const runButton = html.getElementById("runButton");
const textField = html.getElementById("textField");
// initialise application
async () => {
    await Office.onReady(info => {
        if (info.host === Office.HostType.Word) {
            html.addEventListener("DOMContentLoaded", () => {
                attempt(main);
            });
        }
    });
};
// error handling
async function attempt(fn) {
    try {
        await fn();
    }
    catch {
        (lg) => console.error(lg);
    }
}
// event triggers
async function main() {
    // click the run button
    runButton.onclick = async () => {
        await Word.run(runScript);
    };
}
// run button script
async function runScript(context) {
    const content = context.document.body;
    if (textField.value === "") {
        content.insertParagraph("Who goes there?", "End");
    }
    else {
        content.insertParagraph("Hello " + textField.value, "End");
    }
    await context.sync();
}

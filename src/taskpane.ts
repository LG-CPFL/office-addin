// global declarations
const html = document;
const runButton = html.getElementById("runButton") as HTMLButtonElement;
const textField = html.getElementById("textField") as HTMLInputElement;

// initialise application
Office.onReady().then(() => {
    html.addEventListener("DOMContentLoaded", () => {
        attempt(main)
    })
})

// error handling
async function attempt(fn:Function) {
    try {
        await fn();
    }
    catch {
        (lg:Error) => console.error(lg);
    }
}

// event triggers
async function main() {
    // click the run button
    runButton.onclick = async () => {
        await Word.run(runScript)
    }
}

// run button script
async function runScript(context:Word.RequestContext) {
    const content = context.document.body;

    if (textField.value === "") {
        content.insertParagraph("Who goes there?", "End")
    } else {
        content.insertParagraph("Hello " + textField.value, "End")
    }

    await context.sync();
}

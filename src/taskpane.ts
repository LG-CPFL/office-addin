// initialise application
Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        const sideloadMsg = document.getElementById("sideload-msg");
        const appBody = document.getElementById("app-body");
        if (sideloadMsg && appBody) {
            sideloadMsg.style.display = "none";
            appBody.style.display = "flex";
            document.addEventListener("DOMContentLoaded",() => {
                attempt(events)
            })
        } else {
            console.error("Elements missing")
        }
    } else {
        console.error("Host invalid")
    }
})

// error handling
async function attempt(fn:Function) {
    try {
        await fn();
    }
    catch (lg) {
        console.error(lg);
    }
}

// create event triggers
function events() {
    // clicking the run button
    const runButton = document.getElementById("run-button");
    if (runButton) {
        runButton.addEventListener("click",() => {
            attempt(main)
        })
    }
}

// execute script
async function main() {
    return Word.run(async (context:Word.RequestContext) => {
        const content = context.document.body;
        content.insertParagraph("Hello There", "End");
        await context.sync();
    })
}
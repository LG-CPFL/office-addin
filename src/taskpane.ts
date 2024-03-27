// initialise application
Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        document.getElementById("sideload-msg").style.display = "none";
        document.getElementById("app-body").style.display = "flex";
        $(document).on("ready",() => {
            attempt(events)
        })
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

// event triggers
async function events() {
    // click the run button
    $("#run-button").on("click", async () => {
        return Word.run((wrc) => main(wrc))
    })
}

// execute script
async function main(context:Word.RequestContext) {
    const content = context.document.body;

    content.insertParagraph("Hello There", "End")
    
    await context.sync();
}
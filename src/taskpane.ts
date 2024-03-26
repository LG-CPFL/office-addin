// initialise application
Office.onReady((info) => {
    $(document).on("ready", () => {
        if (info.host === Office.HostType.Word) {
            attempt(events)
        } else {
            console.error("Host invalid.")
        }
    })
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
    $("#runButton").on("click", async () => {
        await Word.run(main)
    })
}

// execute script
async function main(context:Word.RequestContext) {
    const content = context.document.body;
    let input = $("#textField").val().toString().trim();

    if (input === "") {
        content.insertParagraph("Who goes there?", "End")
    } else {
        content.insertParagraph("Hello " + input, "End")
    }

    await context.sync();
}
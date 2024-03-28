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
function attempt(fn:Function) {
    try {
        fn();
    }
    catch (lg) {
        console.error(lg);
    }
}

// event triggers
function events() {
    // click the run button
    document.addEventListener("click", () => {
        Word.run((wrc) => main(wrc))
    })
}

// execute script
async function main(context:Word.RequestContext) {
    const content = context.document.body;

    content.insertParagraph("Hello There", "End")
    
    await context.sync();
}
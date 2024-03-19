// global declarations
const html = document;
const runButton = html.getElementById("runButton") as HTMLButtonElement;
const textField = html.getElementById("textField") as HTMLInputElement;

// check that office is ready
Office.onReady( () => {
    // check that the html has loaded
    html.addEventListener("DOMContentLoaded", () => {
        // when the button is clicked
        runButton.addEventListener("click", () => {
            await Word.run(main) // run the main function
            .catch((errorMessage:Error) => console.error(errorMessage)) // unless it breaks
        });
    });
});
  
// test function
async function main(context:Word.RequestContext) {
    const content = context.document.body;

    if (textField.value === "") {
        content.insertParagraph("Who goes there?", "End")
    } else {
        content.insertParagraph("Hello " + textField.value, "End")
    }

    await context.sync();

}
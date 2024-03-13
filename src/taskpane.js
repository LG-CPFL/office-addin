// check that office is ready
Office.onReady( () => {
    // check that the document is loaded
    $(document).ready( () => {
        // when button is clicked
        $("#runButton").on("click", function () {
            Word.run(main) // run main function
            .catch(log => console.error(log)) // unless it breaks
        });
    });
});
  
// script goes here
async function main(context) {
    const document = context.document;

    document.body.insertParagraph("Hello World!", "End")
    await context.sync();

}
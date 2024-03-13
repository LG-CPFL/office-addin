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

    document.body.insertParagraph("Hello World", "End")
    await context.sync();

    console.log("output successful.")

}

// console output to HTML (from GPT)
(function() {
    var oldConsoleLog = console.log;
    console.log = function(message) {
        var consoleOutputDiv = document.getElementById('consoleOutput');
        if (consoleOutputDiv) {
            consoleOutputDiv.innerHTML += message + '<br>';
        }
        oldConsoleLog.apply(console, arguments);
    }
})();
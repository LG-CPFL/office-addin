// import dependencies
import * as Office from 'office.js';
import * as Word from 'office.js/word';

// define taskpane html and elements
const html = document;
const runButton = html.getElementById("runButton");
const textField = html.getElementById("textField");

// check that office is ready
Office.onReady( () => {
    // check that the app is loaded
    html.onReady( () => {
        // when button is clicked
        runButton.addEventListener("click", function () {
            Word.run(main) // run main function
            .catch(log => console.error(log)) // unless it breaks
        });
    });
});
  
// script goes here (testing testing)
async function main(context) {
    const document = context.document;
    
    textField.onload(
        document.body.insertParagraph("Hello " + textField.value, "End")
    );
    await context.sync();

}
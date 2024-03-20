## wip
- figure out how to do this
- is just building my own hotdocs viable?
- would need to create something that can parse script built into a template.

potential libraries to add:
core-js
jquery
office-ui-fabric-js (css?)

## snippets
### error catching

// Default helper for invoking an action and handling errors.
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error(error);
  }
}

// can also use an html element to display error logs as user friendly:

HTML
<div id="errorMessage"></div>
JAVASCRIPT
catch (error) => {document.getElementById("errorMessage").textContent = error}

// or could use a dialogue box (see lower down)

### inserting an external file

HTML
<form>
	<input type="file" id="file" />
</form> 

JAVASCRIPT
$("#file").on("change", getBase64); // not a callback!

let externalDocument;

function getBase64() {
  // Retrieve the file and set up an HTML FileReader element.
  const myFile = <HTMLInputElement>document.getElementById("file");
  const reader = new FileReader();

  reader.onload = (event) => {
    // Remove the metadata before the Base64-encoded string.
    const startIndex = reader.result.toString().indexOf("base64,");
    externalDocument = reader.result.toString().substr(startIndex + 7);
  };

  // Read the file as a data URL so that we can parse the Base64-encoded string.
  reader.readAsDataURL(myFile.files[0]);
}

### pop up message box

HTML
<div id="dialog" style="display: none;">
	<p>This is a custom dialog box!</p>
	<button id="closeButton">Close</button>
</div>

CSS
#dialog {
    position: absolute;
    top: 50%;
    left: 50%;
    transform: translate(-50%, -50%);
    background-color: white;
    padding: 20px;
    border: 5px dotted rgb(194, 89, 28);
    box-shadow: 0 2px 4px rgba(0,0,0,0.2);
}

JAVASCRIPT
// Function to display the dialog box
function showDialog() {
  document.getElementById("dialog").style.display = "block";
}

// Function to close the dialog box
function closeDialog() {
  document.getElementById("dialog").style.display = "none";
}

// Add event listener to close button
document.getElementById("closeButton").addEventListener("click", closeDialog);

// use showDialog() to then activate the pop up.
showDialog();
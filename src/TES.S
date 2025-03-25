$("#run").on("click", () => tryCatch(run));

function run() {
  return Word.run(function(context) {
    var range = context.document.getSelection();
    range.font.color = "red";
    range.load("text");
    replaceID();

    return context.sync().then(function() {
      console.log('The selected text was "' + range.text + '".');
    });
  });
}

function replaceID() {
  return Word.run(async (context) => {
    // Load custom properties
    const metadataID = context.document.properties.customProperties;
    metadataID.load("key, value");
    await context.sync();

    // // Find the custom property with key "VD_ID"
    // const resultId = metadataID.items.find(meta => meta.key === "DocID");

    // // Use a placeholder value for testing if the custom property is not found
    // const replacementText = resultId ? resultId.value : "TESTED";

    // // Search for {{ID}} in the document and replace it with the replacement text
    // const searchResults = context.document.body.search("{{ID}}", { matchCase: true, matchWholeWord: true });
    // context.load(searchResults, "text");
    await context.sync();

    // if (searchResults.items.length > 0) {
    //   searchResults.items.forEach(item => {
    //     item.insertText(replacementText, Word.InsertLocation.replace);
    //   });
    // } else {
    //   console.log("The placeholder '{{ID}}' was not found in the document.");
    //   console.log(resultId);
    // }

    // Display all custom properties in the console for debugging
    // document.getElementById("test").innerHTML = "";
    metadataID.items.forEach(item => {
      console.log(`Custom property: Key = ${item.key}, Value = ${item.value}`);
      // document.getElementById("test").innerHTML += `Key: ${item.key}, Value: ${item.value}<br>`;
    });

    await context.sync();
  });
}
// Default helper for invoking an action and handling errors.
function tryCatch(callback) {
  Promise.resolve()
    .then(callback)
    .catch(function(error) {
      // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
      console.error(error);
    });
}


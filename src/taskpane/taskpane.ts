/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  return Word.run(async (context) => {
    // Load custom properties
    const metadataID = context.document.properties.customProperties;
    metadataID.load("key, value");
    await context.sync();

    // Find the custom property with key "VD_ID"
    const resultId = metadataID.items.find(meta => meta.key === "VD_ID");

    // Use a placeholder value for testing if the custom property is not found
    const replacementText = resultId ? resultId.value : "TESTED";

    // Search for {{ID}} in the document and replace it with the replacement text
    const searchResults = context.document.body.search("{{ID}}", {matchCase: true, matchWholeWord: true});
    context.load(searchResults, "text");
    await context.sync();

    if (searchResults.items.length > 0) {
      searchResults.items.forEach(item => {
        item.insertText(replacementText, Word.InsertLocation.replace);
      });
    } else {
      console.log("The placeholder '{{ID}}' was not found in the document.");
    }

    // Display all custom properties in the console for debugging
    document.getElementById("test").innerHTML = "";
    metadataID.items.forEach(item => {
      console.log(`Custom property: Key = ${item.key}, Value = ${item.value}`);
      document.getElementById("test").innerHTML += `Key: ${item.key}, Value: ${item.value}<br>`;
    });

    await context.sync();
  });
}

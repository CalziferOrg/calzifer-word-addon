/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
import { base64Image } from "../../base64Image";

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

        // Determine if the user's version of Office supports all the Office.js APIs that are used in the tutorial.
    if (!Office.context.requirements.isSetSupported('WordApi', '1.3')) {
      console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
    }

    // Assign event handlers and other initialization logic.
    document.getElementById("insert-paragraph").onclick = insertParagraph;

    document.getElementById("apply-style").onclick = applyStyle;

    document.getElementById("apply-custom-style").onclick = applyCustomStyle;

    document.getElementById("change-font").onclick = changeFont;
    
    document.getElementById("insert-text-into-range").onclick = insertTextIntoRange;

    document.getElementById("insert-text-outside-range").onclick = insertTextBeforeRange;

    document.getElementById("replace-text").onclick = replaceText;

    document.getElementById("insert-image").onclick = insertImage;

    document.getElementById("insert-html").onclick = insertHTML;

    document.getElementById("insert-table").onclick = insertTable;

  }
});

async function insertParagraph() {
  await Word.run(async (context) => {

      // TODO1: Queue commands to insert a paragraph into the document.
      const docBody = context.document.body;
      docBody.insertParagraph("Calzifer Word Addon adds some text to this Word document.", "Start");

      await context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

async function applyStyle() {
  await Word.run(async (context) => {

      // TODO1: Queue commands to style text.
      const firstParagraph = context.document.body.paragraphs.getFirst();
      firstParagraph.styleBuiltIn = Word.Style.intenseReference;

      await context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

async function applyCustomStyle() {
  await Word.run(async (context) => {

      // TODO1: Queue commands to apply the custom style.
      const lastParagraph = context.document.body.paragraphs.getLast();
      lastParagraph.style = "MyCustomStyle";

      await context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

async function changeFont() {
  await Word.run(async (context) => {

      // TODO1: Queue commands to apply a different font.
      const secondParagraph = context.document.body.paragraphs.getFirst().getNext();
      secondParagraph.font.set({
        name: "Times New Roman",
        bold: true,
        size: 14
      });

      await context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

async function insertTextIntoRange() {
  await Word.run(async (context) => {

      const doc = context.document;
      const originalRange = doc.getSelection();
      originalRange.insertText(" (C2R)", "End");

      originalRange.load("text");
      await context.sync();

      doc.body.insertParagraph("Original range: " + originalRange.text, "End");

      await context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

async function insertTextBeforeRange() {
  await Word.run(async (context) => {

      // TODO1: Queue commands to insert a new range before the
      //        selected range.
      const doc = context.document;
      const originalRange = doc.getSelection();
      originalRange.insertText("Microsoft Office, ", "Before");

      // TODO2: Load the text of the original range and sync so that the
      //        range text can be read and inserted.
      originalRange.load("text");
      await context.sync();

      doc.body.insertParagraph("Current text of original range: " + originalRange.text, "End");
      await context.sync();

  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

async function replaceText() {
  await Word.run(async (context) => {

      // TODO1: Queue commands to replace the text.
      const doc = context.document;
      const originalRange = doc.getSelection();
      originalRange.insertText("many", "Replace");

      await context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

async function insertImage() {
  await Word.run(async (context) => {

      // TODO1: Queue commands to insert an image.
      context.document.body.insertInlinePictureFromBase64(base64Image, "End");

      await context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

async function insertHTML() {
  await Word.run(async (context) => {

      // TODO1: Queue commands to insert a string of HTML.
      const blankParagraph = context.document.body.paragraphs.getLast().insertParagraph("", "After");
      blankParagraph.insertHtml('<p style="font-family: verdana;">Inserted HTML.</p><p>Another paragraph</p>', "End");

      await context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}

async function insertTable() {
  await Word.run(async (context) => {

      // TODO1: Queue commands to get a reference to the paragraph
      //        that will proceed the table.
      const secondParagraph = context.document.body.paragraphs.getFirst().getNext();

      // TODO2: Queue commands to create a table and populate it with data.
      const tableData = [
          ["Name", "ID", "Birth City"],
          ["Bob", "434", "Chicago"],
          ["Sue", "719", "Havana"],
      ];
      secondParagraph.insertTable(3, 3, "After", tableData);

      await context.sync();
  })
  .catch(function (error) {
      console.log("Error: " + error);
      if (error instanceof OfficeExtension.Error) {
          console.log("Debug info: " + JSON.stringify(error.debugInfo));
      }
  });
}
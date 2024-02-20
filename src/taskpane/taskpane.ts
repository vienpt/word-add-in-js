/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

let count = 0;

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    //    document.getElementById("insert-paragraph").onclick = () => tryCatch(insertParagraph);
    //    document.getElementById("apply-style").onclick = () => tryCatch(applyStyle);
    //    document.getElementById("apply-custom-style").onclick = () => tryCatch(customStyle);

    document.getElementById("get-content").onclick = () => tryCatch(getContent);

    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    // updateCount();
    // Office.addin.onVisibilityModeChanged(function (args) {
    //   if (args.visibilityMode === "Taskpane") {
    //     updateCount();
    //   }
    // });
  }
});

async function getContent() {
  await Word.run(async (context) => {
    // TODO 1: retrieve the entire text content of the active Word document.
    const body = context.document.body;
    context.load(body);

    await context.sync();

    document.getElementById("textarea-text").textContent = body.text;
  });
}

function updateCount() {
  count++;
  document.getElementById("app-body").textContent = `Task pane opened ${count} times`;
}

async function insertParagraph() {
  await Word.run(async (context) => {
    // TODO: queue commands to insert paragraph
    const docBody = context.document.body;
    docBody.insertParagraph(
      "Office has several versions, including Office 2016, Microsoft 365 subscription, and Office on the web.",
      Word.InsertLocation.start
    );

    await context.sync();
  });
}

async function applyStyle() {
  await Word.run(async (context) => {
    // TODO2: queue commands to apply style
    const firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.styleBuiltIn = Word.BuiltInStyleName.intenseQuote;

    await context.sync();
  });
}

async function customStyle() {
  await Word.run(async (context) => {
    const lastParagraph = context.document.body.paragraphs.getFirst().getNext();
    lastParagraph.font.set({
      name: "Courier New",
      bold: true,
      size: 18,
    });

    await context.sync();
  });
}

async function tryCatch(callback: () => Promise<void>) {
  try {
    await callback();
  } catch (error) {
    throw new Error(error);
  }
}

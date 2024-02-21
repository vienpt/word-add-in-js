/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

import "./taskpane.css";

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    // document.getElementById("get-content").onclick = () => tryCatch(getContent);

    // editor
    // document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
  }
});

async function getTextAreaContent() {
  const textarea = document.getElementById("textarea-text") as HTMLTextAreaElement
  return await selectTextContent(textarea)
}

async function selectTextContent(textarea: HTMLTextAreaElement) {
  const start = textarea.selectionStart;
  const end = textarea.selectionEnd;

  // Get the selected text
  return textarea.value.substring(start, end);
}

async function getContent() {
  await Word.run(async (context) => {
    // TODO 1: retrieve the entire text content of the active Word document.
    const body = context.document.body;
    context.load(body);

    await context.sync();

    document.getElementById("textarea-text").textContent = body.text;
  });
}

async function tryCatch(callback: () => Promise<void>) {
  try {
    await callback();
  } catch (error) {
    throw new Error(error);
  }
}

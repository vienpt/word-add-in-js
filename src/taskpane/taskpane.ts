/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

import "./taskpane.css";

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("btn-review").onclick = () => tryCatch(review);
    document.getElementById("btn-comment").onclick = () => tryCatch(addComment);
    document.getElementById("input-comment").onkeyup = (ev) => tryCatch(() => enterKeyComment(ev));

    document.getElementById("btn-tab-review").onclick = () => tryCatch(setActive("review"))
    document.getElementById("btn-tab-draft").onclick = () => tryCatch(setActive("draft"))

    document.getElementById("app-body").style.display = "flex";
  }
});

async function setActive(tab: string) {
  if (tab === "review") {
    document.getElementById("btn-tab-review").classList.add("btn-active");
    document.getElementById("btn-tab-draft").classList.remove("btn-active");

    document.getElementById("tab-review").style.display = "flex";
    document.getElementById("tab-draft").style.display = "none";
  } else {
    document.getElementById("btn-tab-draft").classList.add("btn-active");
    document.getElementById("btn-tab-review").classList.remove("btn-active");

    document.getElementById("tab-draft").style.display = "flex";
    document.getElementById("tab-review").style.display = "none";
  }
}

async function enterKeyComment(ev: KeyboardEvent) {
  document.getElementById("comment-tip").style.display = "flex";
  if (ev.key === "Meta") {
    ev.preventDefault();
    // Trigger the button element with a click
    await addComment();
  }
}

async function addComment() {
  // TODO: get input comment and apply to list-comment
  const commentBox = document.getElementById("input-comment") as HTMLInputElement;
  if (!commentBox.value) {
    return
  }

  const comments = document.getElementById("comments");
  const li = document.createElement("li");

  comments.style.display = "flex";
  li.appendChild(document.createTextNode(commentBox.value));
  li.setAttribute("class", "comment-item");
  comments.appendChild(li);

  commentBox.value = '';
}

async function review() {
  // TODO: get new edited text in textarea and apply change to suggestion card
  const content = await getTextAreaContent();
  if (!content) {
    return
  }

  document.getElementById("show-items").style.display = "flex";
  document.getElementById("edited-text").innerHTML = content;
}

async function getTextAreaContent() {
  const textarea = document.getElementById("textarea-text") as HTMLTextAreaElement
  // return await selectTextContent(textarea)
  return textarea.value
}


// async function selectTextContent(textarea: HTMLTextAreaElement) {
//   const start = textarea.selectionStart;
//   const end = textarea.selectionEnd;
//
//   // Get the selected text
//   return textarea.value.substring(start, end);
// }

// async function getContent() {
//   await Word.run(async (context) => {
//     // TODO 1: retrieve the entire text content of the active Word document.
//     const body = context.document.body;
//     context.load(body);
//
//     await context.sync();
//
//     document.getElementById("textarea-text").textContent = body.text;
//   });
// }

async function tryCatch(callback: any) {
  try {
    await callback();
  } catch (error) {
    throw new Error(error);
  }
}

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
  await Word.run(async (context) => {
    const formData = {
      subject: document.getElementById("subject").value,
      comment: document.getElementById("comment").value,
    };

    const docBody = context.document.body;
    docBody.insertParagraph(`Subject: ${formData.subject}`, Word.InsertLocation.end);
    docBody.insertParagraph(`Comment: ${formData.comment}`, Word.InsertLocation.end);

    await context.sync();
  }).catch((error) => {
    console.error("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.error("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

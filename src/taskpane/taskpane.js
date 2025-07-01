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
    const subject = document.getElementById("subject").value;
    const comment = document.getElementById("comment").value;

    const now = new Date();
    const monthLabel = now.toLocaleString('default', { month: 'long', year: 'numeric' }); // Ex: "juillet 2025"
    const dateLabel = now.toLocaleDateString(); // Ex: "01/07/2025"

    const body = context.document.body;
    body.load("paragraphs");
    await context.sync();

    const paragraphs = body.paragraphs.items.map(p => p.text.trim());

    // Vérifie si le mois existe déjà
    let monthExists = paragraphs.some(p => p.toLowerCase() === monthLabel.toLowerCase());
    if (!monthExists) {
      body.insertParagraph(monthLabel, Word.InsertLocation.end).font.set({ bold: true, size: 18 });
    }

    // Vérifie si le jour existe déjà
    let dateExists = paragraphs.some(p => p === dateLabel);
    if (!dateExists) {
      body.insertParagraph(dateLabel, Word.InsertLocation.end).font.set({ bold: true, size: 14 });
    }

    // Ajoute les données sous le jour
    body.insertParagraph(`Subject: ${subject}`, Word.InsertLocation.end);
    body.insertParagraph(`Comment: ${comment}`, Word.InsertLocation.end);
    body.insertParagraph("", Word.InsertLocation.end); // saut de ligne

    await context.sync();
  }).catch((error) => {
    console.error("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.error("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

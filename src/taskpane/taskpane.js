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
    const timeLabel = now.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit' }); // Ex: "14:23"

    const body = context.document.body;
    body.load("paragraphs");
    await context.sync();

    // Find the month paragraph (from top)
    const paragraphs = body.paragraphs.items;
    let monthIndex = paragraphs.findIndex(p => p.text.trim().replace(/^🟦\s*/, '').toLowerCase() === monthLabel.toLowerCase());

    // If month doesn't exist, insert at top and update monthIndex
    if (monthIndex === -1) {
      body.insertParagraph(`🟦 ${monthLabel}`, Word.InsertLocation.start).font.set({ bold: true, size: 18 });
      await context.sync();
      body.load("paragraphs");
      await context.sync();
      monthIndex = 0;
    }

    // Find the date paragraph under the month
    const updatedParagraphs = body.paragraphs.items;
    let dateIndex = -1;
    for (let i = monthIndex + 1; i < updatedParagraphs.length; i++) {
      const text = updatedParagraphs[i].text.trim();
      if (text.startsWith("🟦")) break; // Next month section
      if (text.replace(/^🔹\s*/, '') === dateLabel) {
      dateIndex = i;
      break;
      }
    }

    // If date doesn't exist, insert after month
    if (dateIndex === -1) {
      updatedParagraphs[monthIndex].insertParagraph(`   🔹 ${dateLabel}`, Word.InsertLocation.after).font.set({ bold: true, size: 14 });
      await context.sync();
      body.load("paragraphs");
      await context.sync();
      // Find the new dateIndex
      const refreshedParagraphs = body.paragraphs.items;
      for (let i = monthIndex + 1; i < refreshedParagraphs.length; i++) {
      const text = refreshedParagraphs[i].text.trim();
      if (text.startsWith("🟦")) break;
      if (text.replace(/^🔹\s*/, '') === dateLabel) {
        dateIndex = i;
        break;
      }
      }
    }

    // Insert entry after date (so most recent is always just under date)
    const entry = `       ⏰ ${timeLabel} - ${subject}`;
    const commentEntry = `                ${comment}`;
    context.document.body.insertParagraph("", Word.InsertLocation.end);
    const inserted = body.paragraphs.items[dateIndex].insertParagraph(entry, Word.InsertLocation.after);
    const commentLine = inserted.insertParagraph(commentEntry, Word.InsertLocation.after);
    commentLine.insertParagraph("", Word.InsertLocation.after);


    await context.sync();
  }).catch((error) => {
    console.error("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.error("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}
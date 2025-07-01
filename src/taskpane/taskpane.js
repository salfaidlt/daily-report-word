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
    const dateString = now.toLocaleDateString();
    const timeString = now.toLocaleTimeString();

    // Lire les données existantes dans le document
    const tables = context.document.body.tables;
    tables.load("text");
    await context.sync();

    let tableData = [];
    
    if (tables.items.length > 0) {
      const existingTable = tables.items[0];
      const existingData = existingTable.getRange().split(["\n"], null, true);
      existingData.load("text");
      await context.sync();

      // Convertir les données existantes en tableau
      const existingRows = existingData.items.map(row => row.text.split("\t"));
      tableData = existingRows;
    }

    // Ajouter les nouvelles données
    tableData.unshift([subject, comment, dateString, timeString]);

    // Trier les données par date
    tableData.sort((a, b) => {
      const dateA = new Date(a[2]);
      const dateB = new Date(b[2]);
      return dateB - dateA;
    });

    // Insérer les données triées dans le document
    const table = context.document.body.insertTable(tableData.length, 4, Word.InsertLocation.start, tableData);
    table.styleBuiltIn = Word.Style.gridTable5Dark_Accent1;
    table.font.color = "black";

    // Ajouter des séparateurs ou des couleurs pour les différents jours et mois
    for (let i = 1; i < tableData.length; i++) {
      const currentDate = new Date(tableData[i][2]);
      const previousDate = new Date(tableData[i - 1][2]);

      if (currentDate.getDate() !== previousDate.getDate()) {
        // Ajouter un séparateur pour les jours différents
        const separatorRow = table.rows.getItemAt(i - 1);
        separatorRow.font.highlightColor = "yellow";
      }

      if (currentDate.getMonth() !== previousDate.getMonth()) {
        // Ajouter un séparateur pour les mois différents
        const separatorRow = table.rows.getItemAt(i - 1);
        separatorRow.font.highlightColor = "green";
      }
    }

    await context.sync();
  });
}

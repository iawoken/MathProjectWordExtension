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
  $.ajax({
    url: "https://stofy.dikholding.com/formulaAPI.php?id=1",
    success: function(response) {
      try {
        const x = response
        x.forEach(async (formula) => {
          await insertFormulaIntoWord(`${formula.formula_title}: ${formula.formula}`);
        });
      } catch (e) {
        console.error('JSON Parse Error:', e);
      }
    },
    error: function(xhr, status, error) {
      console.error(status, error);
    }
  });
}

async function insertFormulaIntoWord(formula) {
  await Word.run(async (context) => {
      const body = context.document.body;
      const xd = `<span id="main-input" id="src1" name="maininput" class="mathquill-input mathquill-editable" rel="tooltip" title>${formula}</span>`
      body.insertHtml(xd, Word.InsertLocation.end);
      await context.sync();
  });
}
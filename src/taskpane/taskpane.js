/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = insertFormulaImages;
  }
});

export async function insertFormulaImages() {
  try {
    const response = await fetch("https://stofy.dikholding.com/formulaAPI.php?id=" + document.getElementById("userid").value);
    if (response.status !== 200) {
      document.getElementById("awokenmsg").innerHTML = "HATA: KULLANICI SISTEMDE KAYITLI DEGIL! <br><br>";
      throw new Error(`KULLANICI SISTEMDE KAYITLI DEGIL!`);
    }

    const formulas = await response.json();
    for (const formula of formulas) {
      const imagePath = await convertFormulaToImage(formula.formula);
      await insertImageIntoWord(imagePath);
    }
  } catch (e) {
    console.error('Error:', e);
  }
}

function blobToBase64(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => resolve(reader.result.split(',')[1]);
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

async function convertFormulaToImage(formula) {
  const response = await fetch("https://latex.codecogs.com/png.latex?" + encodeURIComponent(formula));
  const imageBlob = await response.blob();
  const base64Image = await blobToBase64(imageBlob);
  return base64Image;
}

async function insertImageIntoWord(base64Image) {
  await Word.run(async (context) => {
    const body = context.document.body;
    const range = body.getRange();
    range.insertInlinePictureFromBase64(base64Image, Word.InsertLocation.end);
    await context.sync();
  });
}
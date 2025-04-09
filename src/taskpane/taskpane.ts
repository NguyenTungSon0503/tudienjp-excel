/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("app-body").style.display = "flex";
    Excel.run(async (context) => {
      const workbook = context.workbook;
      workbook.onSelectionChanged.add(handleSelectionChanged);
      await context.sync();
    });
  }
});

async function handleSelectionChanged(eventArgs) {
  const element = document.getElementById("tudienjp").childNodes[1] as HTMLElement;
  if (element) {
    element.style.display = "none";
  }
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();

      range.load(["address", "values", "format/font/name"]);

      await context.sync();
      const value = range.values[0][0];

      // Regex to check if the value contains Japanese characters
      const japaneseRegex = /[\u3040-\u30FF\u4E00-\u9FFF]/;
      const isJapanese = japaneseRegex.test(value);

      document.getElementById("cell-value").style.fontFamily =
        range.format.font.name ?? "游ゴシック";

      document.getElementById("cell-value").innerText = isJapanese ? value : "";
    });
  } catch (error) {
    console.error("Error in selection change handler:", error);
  }
}

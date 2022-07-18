/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

import {
  provideFluentDesignSystem,
  fluentDialog, 
  fluentButton,
} from "@fluentui/web-components";

provideFluentDesignSystem().register(fluentButton(), fluentDialog());

let dialogElement;
dialogElement = document.getElementById('defaultDialog');

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("message").innerHTML = "Ready to receive files.";
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = runOnclick;
  }
});

const getDataFromSelection  = async (context: Excel.RequestContext) => {
  if(Office.context.document.getSelectedDataAsync){

  }
}

const paintCells = async (context: Excel.RequestContext) => {
  const range = context.workbook.getSelectedRange();
  range.format.fill.color = "red";
  await context.sync();
}

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    Excel.run((context)=>{
      paintCells(context);   
      return context.sync().then(()=>{
        
      });
    })
  }
});

export async function runOnclick() {
  try {
    await Excel.run(async (context) => {
      dialogElement.hidden = false;
      const range = context.workbook.getSelectedRange();
      range.load("address");
      range.format.fill.color = "yellow";
      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}

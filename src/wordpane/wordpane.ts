
Office.onReady(async (info) => {
  debugger;
  if (info.host === Office.HostType.Word) {
    document.getElementById("btnPaintText").onclick = onPaintText;
    document.getElementById("btnReplaceText").onclick = onReplaceAllText;
  }
});

export async function onPaintText(){
  try{
    await Word.run(async (context) => {
      const range = context.document.getSelection();
      range.load("text");
      range.font.color = "red";
      await context.sync();
      console.log("selected text :"+range.text);
    })
  }catch(e){
    console.log("error on paint text");
  }
}

export async function onReplaceAllText(){
  try{
    await Word.run(async (context) => {
      let body = context.document.body;
      body.insertText("", "Replace");
      body.insertParagraph("Template document", "Start").styleBuiltIn = Word.Style.heading1;
      body.insertParagraph("Here is a paragraph", "End").styleBuiltIn = Word.Style.normal;
      body.insertParagraph("Here is a another p", "End").styleBuiltIn = Word.Style.normal;
      await context.sync();
    })
  }catch(e){
    console.log("error on insert text");
  }
}


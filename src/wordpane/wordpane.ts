
Office.onReady(async (info) => {
  debugger;
  if (info.host === Office.HostType.Word) {
    document.getElementById("btnPaintText").onclick = onPaintText;
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
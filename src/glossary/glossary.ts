
Office.onReady(async (info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("btnPaintText").onclick = onPaintText;
    document.getElementById("btnReplaceText").onclick = onReplaceAllText;
    document.getElementById("insertControls").onclick = onInsertControls;
    document.getElementById("setup").onclick = setup;
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

export async function onInsertControls(){
  try{
    await Word.run(async (context) => {
      let paragraphs = context.document.body.paragraphs;
      paragraphs.load("$none"); // No properties needed.

      await context.sync();

      const list = paragraphs.items[1].startNewList(); // Indicates new list to be started in the second paragraph.
      list.load("$none"); // No properties needed.

      await context.sync();

      // To add new items to the list use start/end on the insert location parameter.
      list.insertParagraph("New list item on top of the list", "Start");
      let paragraph = list.insertParagraph("New list item at the end of the list (4th level)", "End");
      paragraph.listItem.level = 4; // Sets up list level for the list item.
      // To add paragraphs outside the list use before/after:
      list.insertParagraph("New paragraph goes after (not part of the list)", "After");

      await context.sync();
    });
  }catch(e){
    console.log(e);
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


async function setup() {
  await Word.run(async (context) => {
    const body = context.document.body;
    body.clear();
    body.insertParagraph(
      "Themes and styles also help keep your document coordinated. When you click design and choose a new Theme, the pictures, charts, and SmartArt graphics change to match your new theme. When you apply styles, your headings change to match the new theme. ",
      "Start"
    );
    body.insertParagraph(
      "Save time in Word with new buttons that show up where you need them. To change the way a picture fits in your document, click it and a button for layout options appears next to it. When you work on a table, click where you want to add a row or a column, and then click the plus sign. ",
      "Start"
    );
    body.insertParagraph(
      "Video provides a powerful way to help you prove your point. When you click Online Video, you can paste in the embed code for the video you want to add. You can also type a keyword to search online for the video that best fits your document.",
      "Start"
    );
    body.paragraphs
      .getLast()
      .insertText(
        "To make your document look professionally produced, Word provides header, footer, cover page, and text box designs that complement each other. For example, you can add a matching cover page, header, and sidebar. Click Insert and then choose the elements you want from the different galleries. ",
        "Replace"
      );
  });
}

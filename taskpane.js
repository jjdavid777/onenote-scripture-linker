Office.onReady(() => {
  console.log("Scripture Linker Add-in loaded");
  console.log("Office.js is ready in OneNote");

  document.getElementById("linkButton").addEventListener("click", async () => {
    try {
      await OneNote.run(async context => {
        const page = context.application.getActivePage();
        const selection = context.application.getSelectedText();
        selection.load("text");
        await context.sync();

        const reference = selection.text.trim();
        if (reference) {
          const encodedRef = encodeURIComponent(reference);
          const url = `https://www.bible.com/bible/116/${encodedRef}`;
          const paragraph = page.addRichText(`<a href="${url}">${reference}</a>`);
          await context.sync();
          console.log("Scripture linked:", url);
        } else {
          console.log("No text selected.");
        }
      });
    } catch (error) {
      console.error("Error linking scripture:", error);
    }
  });
});

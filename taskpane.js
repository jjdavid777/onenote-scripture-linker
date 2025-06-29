
Office.onReady(() => {
  console.log("Scripture Linker Add-in loaded");

  document.getElementById("linkButton").onclick = async function linkScriptures() {
    try {
      await OneNote.run(async context => {
        const page = context.application.getActivePage();
        const paragraphs = page.contents;
        context.load(paragraphs, "id,paragraph/text");

        await context.sync();

        const scriptureRegex = /\b(?:[1-3]\s)?[A-Z][a-z]+\s\d{1,3}:\d{1,3}\b/g;

        for (let i = 0; i < paragraphs.items.length; i++) {
          const paragraph = paragraphs.items[i].paragraph;
          if (paragraph && paragraph.text) {
            const matches = paragraph.text.match(scriptureRegex);
            if (matches) {
              for (const match of matches) {
                const encodedRef = encodeURIComponent(match);
                const url = "https://www.bible.com/bible/116/" + encodedRef;
                paragraph.insertHtml(`<a href='${url}'>${match}</a>`, "Replace");
              }
            }
          }
        }

        await context.sync();
      });
    } catch (error) {
      console.error("Error linking scripture:", error);
    }
  };
});

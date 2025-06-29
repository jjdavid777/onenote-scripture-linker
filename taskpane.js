Office.onReady(() => {
  console.log("Office.js is ready in OneNote");
  document.getElementById("linkScriptures").addEventListener("click", linkScriptures);
});

async function linkScriptures() {
  try {
    await OneNote.run(async context => {
      const pages = context.application.getActivePage();
      pages.load("id,contents");
      await context.sync();

      const page = pages;
      const outline = page.contents.items[0].outline;
      outline.load("id,paragraphs");
      await context.sync();

      const paragraphs = outline.paragraphs;
      paragraphs.load("items/id,richText/text");
      await context.sync();

      const scriptureRegex = /\b(?:[1-3]\s)?[A-Z][a-z]+\s\d{1,3}:\d{1,3}\b/g;

      for (let i = 0; i < paragraphs.items.length; i++) {
        const para = paragraphs.items[i];
        const text = para.richText.text;
        const matches = text.match(scriptureRegex);
        if (matches) {
          let newText = text;
          matches.forEach(ref => {
            const encodedRef = encodeURIComponent(ref);
            const url = `https://www.bible.com/bible/116/${encodedRef}`;
            const link = `<a href='${url}' target='_blank'>${ref}</a>`;
            newText = newText.replace(ref, link);
          });
          para.richText.insertHtml(newText, "Replace");
        }
      }

      await context.sync();
      document.getElementById("status").innerText = "Scripture references linked!";
    });
  } catch (error) {
    console.error("Error linking scripture:", error);
    document.getElementById("status").innerText = "Error linking scripture references.";
  }
}

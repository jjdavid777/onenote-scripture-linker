
Office.onReady(() => {
  console.log("Scripture Linker Add-in loaded");
  if (Office.context.host === Office.HostType.OneNote) {
    console.log("Office.js is ready in OneNote");
    document.getElementById("linkScriptures").onclick = linkScriptures;
  }
});

async function linkScriptures() {
  try {
    await OneNote.run(async context => {
      const page = context.application.getActivePage();
      const outlines = page.contents;
      context.load(outlines, "id,type,outline");

      await context.sync();

      const scriptureRegex = /\b(?:[1-3]\s)?[A-Z][a-z]+\s\d{1,3}:\d{1,3}\b/g;

      for (let i = 0; i < outlines.items.length; i++) {
        const outline = outlines.items[i];
        if (outline.type === "Outline") {
          const paragraphs = outline.outline.paragraphs;
          context.load(paragraphs, "items");

          await context.sync();

          for (let j = 0; j < paragraphs.items.length; j++) {
            const paragraph = paragraphs.items[j];
            const text = paragraph.text;
            const matches = text.match(scriptureRegex);

            if (matches) {
              let updatedText = text;
              matches.forEach(ref => {
                const encodedRef = encodeURIComponent(ref);
                const url = `https://www.bible.com/bible/116/${encodedRef}`;
                const link = `<a href="${url}" target="_blank">${ref}</a>`;
                updatedText = updatedText.replace(ref, link);
              });
              paragraph.insertHtml(updatedText, "Replace");
            }
          }
        }
      }

      await context.sync();
      console.log("Scripture references linked.");
    });
  } catch (error) {
    console.error("Error linking scripture:", error);
  }
}

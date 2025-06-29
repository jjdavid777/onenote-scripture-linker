Office.onReady(() => {
  console.log("Scripture Linker Add-in loaded");
  document.getElementById("linkScriptures").onclick = linkScriptures;
});

async function linkScriptures() {
  try {
    await OneNote.run(async context => {
      console.log("Running inside OneNote Web");

      const page = context.application.getActivePage();
      const outlineCollection = page.contents;
      outlineCollection.load("items/type,paragraph/text");

      await context.sync();

      const scriptureRegex = /\b(?:[1-3]\s)?[A-Z][a-z]+\s\d{1,3}:\d{1,3}(?:[-–]\d{1,3})?\b/g;

      for (let i = 0; i < outlineCollection.items.length; i++) {
        const item = outlineCollection.items[i];
        if (item.type === "Paragraph" && item.paragraph && typeof item.paragraph.text === "string") {
          const text = item.paragraph.text;
          const matches = text.match(scriptureRegex);
          if (matches) {
            let updatedText = text;
            matches.forEach(ref => {
              const encodedRef = encodeURIComponent(ref);
              const url = `https://www.bible.com/bible/116/${encodedRef}`;
              const link = `<a href="${url}" target="_blank">${ref}</a>`;
              updatedText = updatedText.replace(ref, link);
            });
            item.paragraph.insertHtml(updatedText, "Replace");
          }
        } else {
          console.log("Skipping item due to missing or invalid paragraph text:", item);
        }
      }

      await context.sync();
      console.log("Scripture linking complete.");
    });
  } catch (error) {
    console.error("Error linking scripture:", error);
  }
}


Office.onReady(() => {
  console.log("Scripture Linker Add-in loaded");

  document.getElementById("linkScriptureButton").addEventListener("click", linkScriptures);
});

async function linkScriptures() {
  try {
    await OneNote.run(async context => {
      console.log("Running inside OneNote Web");

      const page = context.application.getActivePage();
      const outlineCollection = page.contents;
      outlineCollection.load("items");

      await context.sync();

      const scriptureRegex = /\b(?:[1-3]\s)?[A-Z][a-z]+\s\d{1,3}:\d{1,3}(?:[-–]\d{1,3})?\b/g;

      for (let i = 0; i < outlineCollection.items.length; i++) {
        const outline = outlineCollection.items[i];
        const outlineParagraphs = outline.outlineElements;
        outlineParagraphs.load("items/type,paragraph/text");

        await context.sync();

        for (let j = 0; j < outlineParagraphs.items.length; j++) {
          const element = outlineParagraphs.items[j];

          if (element.type === "Paragraph" && element.paragraph && typeof element.paragraph.text === "string") {
            const text = element.paragraph.text;
            console.log(`Processing paragraph [${i}, ${j}]: "${text}"`);

            const matches = text.match(scriptureRegex);
            if (matches) {
              let updatedText = text;
              matches.forEach(ref => {
                const encodedRef = encodeURIComponent(ref);
                const url = `https://www.bible.com/bible/116/${encodedRef}`;
                const link = `<a href="${url}" target="_blank">${ref}</a>`;
                updatedText = updatedText.replace(ref, link);
              });
              element.insertHtml(updatedText, "Replace");
            } else {
              console.log(`No scripture references found in paragraph [${i}, ${j}]`);
            }
          } else {
            console.warn(`Skipping element [${i}, ${j}] due to missing or invalid paragraph text`, element);
          }
        }
      }

      await context.sync();
      console.log("Finished linking scripture references.");
    });
  } catch (error) {
    console.error("Error linking scripture:", error);
  }
}

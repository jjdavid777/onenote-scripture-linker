Office.onReady(() => {
  console.log("Running inside OneNote Web");

  document.getElementById("linkScriptures").addEventListener("click", linkScriptures);
});

async function linkScriptures() {
  try {
    await OneNote.run(async context => {
      const page = context.application.getActivePage();
      const contents = page.contents;
      contents.load("items");
      await context.sync();

      const scriptureRegex = /\b(?:[1-3]\s)?[A-Z][a-z]+\s\d{1,3}:\d{1,3}(?:[-–]\d{1,3})?\b/g;

      for (const item of contents.items) {
        if (item.type === "Outline" && item.outline) {
          const paragraphs = item.outline.paragraphs;
          paragraphs.load("items");
          await context.sync();

          for (const para of paragraphs.items) {
            if (para && typeof para.text === "string") {
              const matches = para.text.match(scriptureRegex);
              if (matches) {
                let updatedText = para.text;
                matches.forEach(ref => {
                  const encodedRef = encodeURIComponent(ref);
                  const url = `https://www.bible.com/bible/116/${encodedRef}`;
                  const link = `<a href="${url}" target="_blank">${ref}</a>`;
                  updatedText = updatedText.replace(ref, link);
                });
                para.insertHtml(updatedText, "Replace");
              }
            } else {
              console.log("Skipping paragraph due to missing or invalid text:", para);
            }
          }
        }
      }

      await context.sync();
      console.log("Scripture linking complete.");
    });
  } catch (error) {
    console.error("Error linking scripture:", error);
  }
}

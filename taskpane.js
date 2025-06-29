
Office.onReady(() => {
  console.log("Scripture Linker Add-in loaded");

  document.getElementById("linkScriptures").onclick = async function linkScriptures() {
    try {
      await OneNote.run(async context => {
        console.log("Office.js is ready in OneNote");
        console.log("Running inside OneNote Web");

        const page = context.application.getActivePage();
        const paragraphs = page.contents;
        context.load(paragraphs, "id,type,paragraph/text");

        await context.sync();

        const scriptureRegex = /\b(?:[1-3]\s)?[A-Z][a-z]+\s\d{1,3}:\d{1,3}\b/g;

        paragraphs.items.forEach(paragraph => {
          if (
            paragraph &&
            paragraph.type === "Paragraph" &&
            paragraph.paragraph &&
            typeof paragraph.paragraph.text === "string"
          ) {
            const matches = paragraph.paragraph.text.match(scriptureRegex);
            if (matches) {
              let updatedText = paragraph.paragraph.text;
              matches.forEach(ref => {
                const encodedRef = encodeURIComponent(ref);
                const url = `https://www.bible.com/bible/116/${encodedRef}`;
                const link = `<a href="${url}" target="_blank">${ref}</a>`;
                updatedText = updatedText.replace(ref, link);
              });
              paragraph.paragraph.insertHtml(updatedText, "Replace");
            }
          } else {
            console.log("Skipping paragraph due to missing or invalid text:", paragraph);
          }
        });

        await context.sync();
        console.log("Scripture references linked.");
      });
    } catch (error) {
      console.error("Error linking scripture:", error);
    }
  };
});

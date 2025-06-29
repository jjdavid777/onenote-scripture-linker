
console.log("Scripture Linker Add-in loaded");

Office.onReady(function(info) {
    if (info.host === Office.HostType.OneNote) {
        console.log("Office.js is ready in OneNote");

        // Add event listener to the document body for keyup events
        document.addEventListener("keyup", function(event) {
            if (event.key === " ") {
                console.log("Space bar pressed");

                // Get the current selection
                Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function(asyncResult) {
                    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                        const selectedText = asyncResult.value.trim();
                        console.log("Selected text:", selectedText);

                        // Match scripture reference pattern (e.g., John 3:16 or Gal 4:13)
                        const scriptureRegex = /^(\w{2,}|[1-3]?\s?[A-Za-z]+)\s(\d+):(\d+)$/;
                        const match = selectedText.match(scriptureRegex);

                        if (match) {
                            const book = match[1].replace(/\s+/g, "+");
                            const chapter = match[2];
                            const verse = match[3];
                            const url = `https://www.bible.com/bible/1/${book}+${chapter}:${verse}`;
                            const linkText = `${match[1]} ${chapter}:${verse}`;

                            console.log("Detected scripture reference:", linkText);
                            console.log("Generated URL:", url);

                            // Replace selected text with hyperlink
                            Office.context.document.setSelectedDataAsync(
                                `<a href='${url}' target='_blank'>${linkText}</a>`,
                                { coercionType: Office.CoercionType.Html },
                                function(result) {
                                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                                        console.log("Hyperlink inserted successfully");
                                    } else {
                                        console.error("Failed to insert hyperlink:", result.error.message);
                                    }
                                }
                            );
                        } else {
                            console.log("No scripture reference detected");
                        }
                    } else {
                        console.error("Failed to get selected text:", asyncResult.error.message);
                    }
                });
            }
        });
    }
});

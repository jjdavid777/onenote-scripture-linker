Office.onReady(() => {
  Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, () => {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        const text = result.value;
        const updatedText = text.replace(/\b(John|Gal|Galatians|Romans|Rom|1\s?Cor|2\s?Cor|1\s?Pet|2\s?Pet|Matt|Matthew|Mark|Luke|Acts|Heb|Hebrews|Rev|Revelation)\s\d{1,3}:\d{1,3}\b/g, (match) => {
          const ref = match.replace(/\s+/g, '+');
          return `<a href="https://www.bible.com/bible/1/${ref}" target="_blank">${match}</a>`;
        });
        Office.context.document.setSelectedDataAsync(updatedText, { coercionType: Office.CoercionType.Html });
      }
    });
  });
});

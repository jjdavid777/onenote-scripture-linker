
console.log("Scripture Linker Add-in loaded");

Office.onReady(() => {
  console.log("Office.js is ready in OneNote");

  // Example: log when the document is ready
  if (Office.context.host === Office.HostType.OneNote) {
    console.log("Running inside OneNote Web");
  }

  // Add a simple event listener for demonstration
  document.addEventListener("keydown", (event) => {
    if (event.code === "Space") {
      console.log("Space bar pressed");
    }
  });
});

// Wait until Office.js is fully ready
Office.onReady(() => {
  // Attach click handler to the button
  document.getElementById("searchButton").onclick = runSearch;
});

// Function to scan body, query API, and update body
async function runSearch() {
  // Get the appointment body text as plain text
  Office.context.mailbox.item.body.getAsync("text", async function (result) {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const body = result.value;
      
      // Use regex to find a line starting with "//search " followed by query text
      const match = body.match(/\/\/search (.+)/i);
      if (!match) {
        showStatus("❌ No '//search' command found in body.");
        return;
      }

      const query = match[1].trim();  // Extract query text
      
      // Call DummyJSON API to search products by query
      const suggestion = await getSuggestion(query);

      // Replace the entire "//search ..." line with the suggestion text
      const newBody = body.replace(/\/\/search .+/, `📌 Suggestion: ${suggestion}`);

      // Update the appointment body with the new content
      Office.context.mailbox.item.body.setAsync(
        newBody,
        { coercionType: Office.CoercionType.Text },
        (res) => {
          if (res.status === Office.AsyncResultStatus.Succeeded) {
            showStatus("✅ Suggestion inserted successfully.");
          } else {
            showStatus("❌ Failed to insert suggestion.");
          }
        }
      );
    } else {
      showStatus("❌ Failed to get appointment body.");
    }
  });
}

// Function to call DummyJSON API and return first product title or fallback text
async function getSuggestion(query) {
  try {
    const res = await fetch(`https://dummyjson.com/products/search?q=${encodeURIComponent(query)}`);
    const data = await res.json();
    return data.products[0]?.title || "No results found";
  } catch (error) {
    return "Error fetching data";
  }
}

// Utility function to update status text in UI
function showStatus(msg) {
  document.getElementById("status").textContent = msg;
}

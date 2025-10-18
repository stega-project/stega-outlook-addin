Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("btnHeaders").onclick = getHeaders;
    loadEmailInfo();
  }
});

function loadEmailInfo() {
  const item = Office.context.mailbox.item;
  document.getElementById("from").textContent = item.from.displayName + " <" + item.from.emailAddress + ">";
  document.getElementById("to").textContent = item.to.map(r => r.displayName).join(", ");
  document.getElementById("subject").textContent = item.subject;
  document.getElementById("date").textContent = item.dateTimeCreated.toLocaleString();
}

async function getHeaders() {
  const item = Office.context.mailbox.item;
  item.getAllInternetHeadersAsync(result => {
    if (result.status === Office.AsyncResultStatus.Succeeded) {
      const headers = result.value;
      document.getElementById("headers").textContent = headers;
    } else {
      document.getElementById("headers").textContent = "Impossible de lire les en-têtes.";
    }
  });
}

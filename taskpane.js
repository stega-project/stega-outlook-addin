let headersCache = null;
let mailboxBootstrapped = false;

function bootstrapMailboxUi() {
  if (mailboxBootstrapped) {
    return;
  }

  mailboxBootstrapped = true;

  const wireHandlers = () => {
    const headersButton = document.getElementById("btnHeaders");
    const copyButton = document.getElementById("btnCopyHeaders");

    if (headersButton) {
      headersButton.addEventListener("click", getHeaders);
    }

    if (copyButton) {
      copyButton.addEventListener("click", copyHeaders);
    }

    loadEmailInfo().catch(error => reportError("Unable to load message information.", error));
  };

  if (document.readyState === "complete" || document.readyState === "interactive") {
    wireHandlers();
  } else {
    document.addEventListener("DOMContentLoaded", wireHandlers, { once: true });
  }
}

Office.onReady(info => {
  if (info.host === Office.HostType.Outlook) {
    bootstrapMailboxUi();
  }
});

if (typeof Office !== "undefined") {
  const previousInitialize = Office.initialize;
  Office.initialize = function initializeOverride(...args) {
    bootstrapMailboxUi();

    if (typeof previousInitialize === "function") {
      previousInitialize.apply(this, args);
    }
  };
}

async function loadEmailInfo() {
  const item = getMailboxItem();
  if (!item) {
    showNotification("Outlook did not provide a message to inspect.", "error");
    updateSignatureStatus({ status: "error", message: "No message context is available." });
    return;
  }

  setText("from", formatSender(item.from));
  setText("to", formatRecipients(item.to));
  setText("subject", item.subject || "—");

  const created = item.dateTimeCreated;
  if (created instanceof Date) {
    setText("date", created.toLocaleString());
  } else if (created) {
    setText("date", new Date(created).toLocaleString());
  } else {
    setText("date", "—");
  }

  updateSignatureStatus(null);
}

function getHeaders() {
  const item = getMailboxItem();
  if (!item) {
    showNotification("No message is available to read headers from.", "error");
    updateSignatureStatus({ status: "error", message: "No message context is available." });
    return;
  }
  const readButton = document.getElementById("btnHeaders");

  readButton.disabled = true;
  setCopyEnabled(false);
  showNotification("Retrieving headers…");
  updateSignatureStatus({ status: "pending", message: "Retrieving headers…" });

  item.getAllInternetHeadersAsync(result => {
    readButton.disabled = false;

    if (result.status === Office.AsyncResultStatus.Succeeded) {
      headersCache = result.value || "";
      document.getElementById("headers").textContent = headersCache || "No headers returned by Outlook.";
      setCopyEnabled(Boolean(headersCache));
      analyzeSignature(headersCache);
      showNotification("Headers retrieved. Scroll down to review them.");
    } else {
      headersCache = null;
      document.getElementById("headers").textContent = "Unable to read the internet headers.";
      updateSignatureStatus({ status: "error", message: "Unable to load headers." });
      showNotification("Outlook could not provide the headers. Try again in a moment.", "error");
    }
  });
}

function analyzeSignature(rawHeaders) {
  if (!rawHeaders) {
    updateSignatureStatus({ status: "warning", message: "No headers returned by Outlook." });
    return;
  }

  const headerMap = parseHeaders(rawHeaders);
  const signature = headerMap["x-stega-signature"] || null;
  const timestamp = headerMap["x-stega-timestamp"] || headerMap["x-stega-date"] || null;
  const verdict = headerMap["x-stega-verdict"] || null;

  if (signature) {
    const normalizedVerdict = verdict ? verdict.trim().toLowerCase() : "";
    let statusVariant = "success";
    let message = "STEGA signature found.";

    if (normalizedVerdict) {
      const validVerdicts = new Set(["valid", "verified", "pass"]);
      const invalidTokens = ["invalid", "fail", "tamper", "revoked"];

      if (invalidTokens.some(token => normalizedVerdict.includes(token))) {
        statusVariant = "error";
        message = `STEGA signature flagged as invalid (${verdict}).`;
      } else if (!validVerdicts.has(normalizedVerdict)) {
        statusVariant = "warning";
        message = `STEGA signature found with verdict: ${verdict}.`;
      }
    }

    updateSignatureStatus({
      status: statusVariant,
      message,
      signature,
      timestamp,
      verdict
    });
  } else {
    updateSignatureStatus({
      status: "warning",
      message: "No STEGA signature present in the headers."
    });
  }
}

function parseHeaders(rawHeaders) {
  const map = {};
  const lines = rawHeaders.split(/\r?\n/);
  let currentKey = null;

  lines.forEach(line => {
    if (!line) {
      currentKey = null;
      return;
    }

    if (/^[\t ]/.test(line) && currentKey) {
      map[currentKey] += " " + line.trim();
      return;
    }

    const separatorIndex = line.indexOf(":");
    if (separatorIndex === -1) {
      currentKey = null;
      return;
    }

    const key = line.slice(0, separatorIndex).trim().toLowerCase();
    const value = line.slice(separatorIndex + 1).trim();
    map[key] = value;
    currentKey = key;
  });

  return map;
}

function copyHeaders() {
  if (!headersCache) {
    return;
  }

  const attemptClipboard = text => {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      return navigator.clipboard.writeText(text);
    }

    return new Promise((resolve, reject) => {
      try {
        const textarea = document.createElement("textarea");
        textarea.value = text;
        textarea.setAttribute("readonly", "");
        textarea.style.position = "absolute";
        textarea.style.left = "-9999px";
        document.body.appendChild(textarea);
        textarea.select();
        document.execCommand("copy");
        document.body.removeChild(textarea);
        resolve();
      } catch (error) {
        reject(error);
      }
    });
  };

  attemptClipboard(headersCache)
    .then(() => {
      showNotification("Headers copied to the clipboard.");
    })
    .catch(error => {
      reportError("Unable to copy the headers to the clipboard.", error);
    });
}

function updateSignatureStatus(details) {
  const statusElement = document.getElementById("signatureStatus");
  const detailsContainer = document.getElementById("signatureDetails");
  const verdictRow = document.getElementById("signatureVerdictRow");

  if (!details) {
    statusElement.textContent = "Waiting for headers…";
    statusElement.className = "status status--pending";
    detailsContainer.hidden = true;
    verdictRow.hidden = true;
    return;
  }

  statusElement.textContent = details.message;
  statusElement.className = `status status--${details.status}`;

  if (details.signature || details.timestamp || details.verdict) {
    detailsContainer.hidden = false;
    document.getElementById("signatureValue").textContent = details.signature || "—";
    document.getElementById("signatureTimestamp").textContent = details.timestamp || "—";

    if (details.verdict) {
      verdictRow.hidden = false;
      document.getElementById("signatureVerdict").textContent = details.verdict;
    } else {
      verdictRow.hidden = true;
    }
  } else {
    detailsContainer.hidden = true;
    verdictRow.hidden = true;
  }
}

function reportError(message, error) {
  console.error(message, error);
  showNotification(message, "error");
}

function getMailboxItem() {
  return Office.context && Office.context.mailbox ? Office.context.mailbox.item : null;
}

function showNotification(message, type = "info") {
  const notification = document.getElementById("notification");

  if (!message) {
    notification.textContent = "";
    notification.className = "notification";
    notification.style.display = "none";
    return;
  }

  notification.textContent = message;
  notification.className = type === "error" ? "notification notification--error" : "notification";
  notification.style.display = "block";
}

function setCopyEnabled(value) {
  document.getElementById("btnCopyHeaders").disabled = !value;
}

function setText(id, text) {
  const element = document.getElementById(id);
  if (element) {
    element.textContent = text;
  }
}

function formatSender(sender) {
  if (!sender) {
    return "—";
  }

  const displayName = sender.displayName && sender.displayName !== sender.emailAddress
    ? sender.displayName
    : null;

  if (displayName && sender.emailAddress) {
    return `${displayName} <${sender.emailAddress}>`;
  }

  return sender.displayName || sender.emailAddress || "—";
}

function formatRecipients(recipients) {
  if (!recipients || !recipients.length) {
    return "—";
  }

  return recipients
    .map(recipient => recipient.displayName || recipient.emailAddress || "Unknown recipient")
    .join(", ");
}

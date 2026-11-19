const form = document.querySelector("#check-form");
const text = document.querySelector("#check-text");
const button = document.querySelector("#check-button");
const result = document.querySelector("#result");
const message = document.querySelector("#message");
const unique = document.querySelector("#unique");
const rows = document.querySelector("#rows");
const copyAllButton = document.querySelector("#copy-all");
const copyHandlesButton = document.querySelector("#copy-handles");
const copyUrlsButton = document.querySelector("#copy-urls");
const pager = document.querySelector("#pager");
const pageLabel = document.querySelector("#page-label");
const pagePrev = document.querySelector("#page-prev");
const pageNext = document.querySelector("#page-next");
const PAGE_SIZE = 10;
let reportRows = [];
let page = 1;

function setBusy(busy) {
  button.disabled = busy;
}

function creatorName(row) {
  const source = `${row.profile_url || ""} ${row.handle || ""} ${row.input || ""}`;
  const match = source.match(/instagram\.com\/([A-Za-z0-9._]+)/i);
  let name = match ? match[1] : String(row.handle || row.input || "");
  name = name.replace(/^@+/, "").replace(/\s+/g, "").toLowerCase();
  name = name.replace(/[^a-z0-9._]/g, "");
  return name;
}

function profileUrl(row) {
  const name = creatorName(row);
  return name ? `https://www.instagram.com/${name}/` : "";
}

function showError(textValue) {
  result.hidden = false;
  message.innerHTML = "";
  const line = document.createElement("p");
  line.className = "invalid";
  line.textContent = textValue;
  message.append(line);
  reportRows = [];
  rows.innerHTML = "";
  unique.hidden = true;
  pager.hidden = true;
}

function renderPage() {
  rows.innerHTML = "";
  const pages = Math.max(1, Math.ceil(reportRows.length / PAGE_SIZE));
  page = Math.min(Math.max(page, 1), pages);
  for (const row of reportRows.slice((page - 1) * PAGE_SIZE, page * PAGE_SIZE)) {
    const item = document.createElement("tr");
    const handle = document.createElement("td");
    handle.textContent = creatorName(row);
    const url = document.createElement("td");
    url.textContent = profileUrl(row);
    item.append(handle, url);
    rows.append(item);
  }
  pager.hidden = reportRows.length <= PAGE_SIZE;
  pageLabel.textContent = `Page ${page} of ${pages}`;
  pagePrev.disabled = page <= 1;
  pageNext.disabled = page >= pages;
}

function showReport(data) {
  result.hidden = false;
  rows.innerHTML = "";
  message.innerHTML = "";
  const added = data.added || 0;
  const exists = data.exists || 0;
  const lines = [];
  if (exists === 1 && !added) {
    lines.push({ status: "exists", text: "Already in the master list" });
  } else if (exists) {
    lines.push({ status: "exists", text: `${exists} already in the master list` });
  }
  if (added === 1) {
    lines.push({ status: "added", text: "1 new, added to the master list" });
  } else if (added) {
    lines.push({ status: "added", text: `${added} new, added to the master list` });
  }
  const reasons = {};
  for (const row of data.rows) {
    if (row.status === "exists" || row.status === "added") continue;
    reasons[row.message] = (reasons[row.message] || 0) + 1;
  }
  for (const [reason, count] of Object.entries(reasons)) {
    lines.push({ status: "invalid", text: `${count} — ${reason}` });
  }
  for (const line of lines) {
    const item = document.createElement("p");
    item.className = line.status;
    item.textContent = line.text;
    message.append(item);
  }
  reportRows = data.rows.filter((row) => row.status === "added");
  unique.hidden = reportRows.length === 0;
  copyAllButton.textContent = "Copy all";
  copyHandlesButton.textContent = "Copy handles";
  copyUrlsButton.textContent = "Copy URLs";
  page = 1;
  renderPage();
}

function uniqueLines(kind) {
  return reportRows.map((row) => {
    const name = creatorName(row);
    const url = profileUrl(row);
    if (kind === "handles") return name;
    if (kind === "urls") return url;
    return `${name}\t${url}`;
  }).filter(Boolean);
}

function copyText(payload) {
  const area = document.createElement("textarea");
  area.value = payload;
  area.style.position = "fixed";
  area.style.top = "0";
  area.style.left = "0";
  area.style.opacity = "0";
  document.body.append(area);
  area.focus();
  area.select();
  area.setSelectionRange(0, area.value.length);
  let copied = false;
  try {
    copied = document.execCommand("copy");
  } catch (error) {
    copied = false;
  }
  area.remove();
  if (!copied && navigator.clipboard && window.isSecureContext) {
    navigator.clipboard.writeText(payload).catch(() => {});
  }
  return copied;
}

function copyUnique(kind, button, label) {
  const payload = uniqueLines(kind).join("\r\n");
  if (!payload) return;
  const copied = copyText(payload);
  button.textContent = copied ? "Copied" : "Copy failed";
  window.setTimeout(() => {
    button.textContent = label;
  }, 1500);
}

async function send(body, isFile) {
  setBusy(true);
  try {
    const response = await fetch("/api/check-list", isFile ? { method: "POST", body } : {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
    const data = await response.json();
    if (!response.ok) throw new Error(data.error || "The list could not be read.");
    showReport(data);
  } catch (error) {
    showError(error.message);
  } finally {
    setBusy(false);
  }
}

form.addEventListener("submit", (event) => {
  event.preventDefault();
  send({ text: text.value }, false);
});

text.addEventListener("keydown", (event) => {
  if (event.key === "Enter" && (event.ctrlKey || event.metaKey)) {
    event.preventDefault();
    form.requestSubmit();
  }
});

copyAllButton.addEventListener("click", () => {
  copyUnique("all", copyAllButton, "Copy all");
});

copyHandlesButton.addEventListener("click", () => {
  copyUnique("handles", copyHandlesButton, "Copy handles");
});

copyUrlsButton.addEventListener("click", () => {
  copyUnique("urls", copyUrlsButton, "Copy URLs");
});

pagePrev.addEventListener("click", () => {
  page -= 1;
  renderPage();
});

pageNext.addEventListener("click", () => {
  page += 1;
  renderPage();
});

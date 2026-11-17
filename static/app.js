const form = document.querySelector("#check-form");
const text = document.querySelector("#check-text");
const button = document.querySelector("#check-button");
const result = document.querySelector("#result");
const message = document.querySelector("#message");
const rows = document.querySelector("#rows");
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

function showError(textValue) {
  result.hidden = false;
  message.innerHTML = "";
  const line = document.createElement("p");
  line.className = "invalid";
  line.textContent = textValue;
  message.append(line);
  reportRows = [];
  rows.innerHTML = "";
  pager.hidden = true;
}

function renderPage() {
  rows.innerHTML = "";
  const pages = Math.max(1, Math.ceil(reportRows.length / PAGE_SIZE));
  page = Math.min(Math.max(page, 1), pages);
  for (const row of reportRows.slice((page - 1) * PAGE_SIZE, page * PAGE_SIZE)) {
    const item = document.createElement("li");
    item.className = row.status;
    const title = document.createElement("strong");
    title.textContent = row.message;
    const name = document.createElement("span");
    name.textContent = row.handle || row.input || "—";
    item.append(title, name);
    if (row.country) {
      const country = document.createElement("span");
      country.textContent = row.country;
      item.append(country);
    }
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
  if (data.rows.length === 1) {
    lines.push({ status: data.rows[0].status, text: data.rows[0].message });
  } else {
    if (exists) lines.push({ status: "exists", text: `${exists} already in the master list` });
    if (added) lines.push({ status: "added", text: `${added} new, added to the master list` });
    const reasons = {};
    for (const row of data.rows) {
      if (row.status === "exists" || row.status === "added") continue;
      reasons[row.message] = (reasons[row.message] || 0) + 1;
    }
    for (const [reason, count] of Object.entries(reasons)) {
      lines.push({ status: "invalid", text: `${count} — ${reason}` });
    }
  }
  for (const line of lines) {
    const item = document.createElement("p");
    item.className = line.status;
    item.textContent = line.text;
    message.append(item);
  }
  reportRows = data.rows;
  page = 1;
  renderPage();
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

pagePrev.addEventListener("click", () => {
  page -= 1;
  renderPage();
});

pageNext.addEventListener("click", () => {
  page += 1;
  renderPage();
});

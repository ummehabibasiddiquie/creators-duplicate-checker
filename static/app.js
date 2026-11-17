const form = document.querySelector("#check-form");
const text = document.querySelector("#check-text");
const fileInput = document.querySelector("#check-file");
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
  fileInput.disabled = busy;
}

function showError(textValue) {
  result.hidden = false;
  message.className = "invalid";
  message.textContent = textValue;
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
  const added = data.added || 0;
  const exists = data.exists || 0;
  if (data.rows.length === 1) {
    message.textContent = data.rows[0].message;
    message.className = data.rows[0].status;
  } else {
    const parts = [];
    if (exists) parts.push(`${exists} already exist`);
    if (added) parts.push(`${added} added to the master list`);
    const other = data.rows.length - exists - added;
    if (other) parts.push(`${other} could not be checked`);
    message.textContent = parts.join(". ") + ".";
    message.className = added && exists ? "mixed" : added ? "added" : exists ? "exists" : "invalid";
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
    if (!response.ok) throw new Error(data.error || "The check could not be finished.");
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

fileInput.addEventListener("change", () => {
  const file = fileInput.files[0];
  fileInput.value = "";
  if (!file) return;
  const body = new FormData();
  body.append("file", file);
  send(body, true);
});

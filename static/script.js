let UPLOAD_ID = null;

const fileInput = document.getElementById("fileInput");
const dealerSelect = document.getElementById("dealerSelect");
const monthSelect = document.getElementById("monthSelect");
const searchBtn = document.getElementById("searchBtn");
const alertSpot = document.getElementById("alertSpot");
const exportPdfBtn = document.getElementById("exportPdfBtn");
const exportExcelBtn = document.getElementById("exportExcelBtn");
const totalItems = document.getElementById("totalItems");
const dealerCodeSpan = document.getElementById("dealerCode");
const monthLabel = document.getElementById("monthLabel");
const summaryRow = document.getElementById("summaryRow");

function showAlert(message, type="success") {
  alertSpot.innerHTML = `<div class="alert alert-${type} alert-dismissible fade show" role="alert">
    ${message}
    <button type="button" class="btn-close" data-bs-dismiss="alert" aria-label="Close"></button>
  </div>`;
}

function clearTable() {
  document.querySelector("#dataTable thead").innerHTML = "";
  document.querySelector("#dataTable tbody").innerHTML = "";
}

function renderTable(rows) {
  clearTable();
  const thead = document.querySelector("#dataTable thead");
  const tbody = document.querySelector("#dataTable tbody");

  if (!rows || rows.length === 0) {
    summaryRow.style.display = "flex";
    totalItems.textContent = 0;
    return;
  }

  const headers = Object.keys(rows[0]);
  const trHead = document.createElement("tr");
  headers.forEach(h => {
    const th = document.createElement("th");
    th.textContent = h;
    trHead.appendChild(th);
  });
  thead.appendChild(trHead);

  rows.forEach(r => {
    const tr = document.createElement("tr");
    headers.forEach(h => {
      const td = document.createElement("td");
      td.textContent = r[h] !== null && r[h] !== undefined ? r[h] : "";
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });
}

fileInput.addEventListener("change", async (e) => {
  const f = e.target.files[0];
  if (!f) return;
  const form = new FormData();
  form.append("file", f);
  const res = await fetch("/upload", { method: "POST", body: form });
  const data = await res.json();
  if (!data.success) {
    showAlert(data.error || "Upload failed", "danger");
    return;
  }
  UPLOAD_ID = data.upload_id;
  showAlert(data.message || "Sheet loaded successfully ✅", "success");

  // Dealers
  dealerSelect.innerHTML = `<option value="" selected>-- Select Dealer --</option>`;
  (data.dealers || []).forEach(d => {
    const opt = document.createElement("option");
    opt.value = d; opt.textContent = d;
    dealerSelect.appendChild(opt);
  });

  // Months (default to ALL)
  monthSelect.innerHTML = `<option value="ALL" selected>ALL Months</option>`;
  (data.months || []).forEach(m => {
    const opt = document.createElement("option");
    opt.value = m; opt.textContent = m;
    monthSelect.appendChild(opt);
  });

  // Reset summary & table
  clearTable();
  totalItems.textContent = 0;
  dealerCodeSpan.textContent = "-";
  monthLabel.textContent = "All";
  summaryRow.style.display = "none";
});

searchBtn.addEventListener("click", async () => {
  if (!UPLOAD_ID) { showAlert("Please upload an Excel sheet first.", "warning"); return; }
  const dealerValue = dealerSelect.value || "";
  const monthValue = monthSelect.value || "ALL";

  const url = new URL("/dealer-data", window.location.origin);
  url.searchParams.set("upload_id", UPLOAD_ID);
  if (dealerValue) url.searchParams.set("dealer", dealerValue);
  if (monthValue) url.searchParams.set("month", monthValue);

  const res = await fetch(url.toString());
  const data = await res.json();
  if (!data.success) {
    showAlert(data.error || "Failed to fetch data", "danger"); return;
  }
  renderTable(data.rows);
  totalItems.textContent = data.total || 0;
  monthLabel.textContent = data.month_label || (monthValue === "ALL" ? "All" : monthValue);
  dealerCodeSpan.textContent = data.dealer_code || "-";
  summaryRow.style.display = "flex";
});

exportExcelBtn.addEventListener("click", () => {
  if (!UPLOAD_ID) { showAlert("Upload a sheet first.", "warning"); return; }
  const dealerValue = dealerSelect.value || "";
  const monthValue = monthSelect.value || "ALL";
  const url = new URL("/export/excel", window.location.origin);
  url.searchParams.set("upload_id", UPLOAD_ID);
  if (dealerValue) url.searchParams.set("dealer", dealerValue);
  if (monthValue) url.searchParams.set("month", monthValue);
  window.location.href = url.toString();
});

exportPdfBtn.addEventListener("click", () => {
  if (!UPLOAD_ID) { showAlert("Upload a sheet first.", "warning"); return; }
  const dealerValue = dealerSelect.value || "";
  const monthValue = monthSelect.value || "ALL";
  const url = new URL("/export/pdf", window.location.origin);
  url.searchParams.set("upload_id", UPLOAD_ID);
  if (dealerValue) url.searchParams.set("dealer", dealerValue);
  if (monthValue) url.searchParams.set("month", monthValue);
  window.location.href = url.toString();
});

function logMessage(text) {
  const box = document.getElementById("chat-output");
  if (!box) return;
  const div = document.createElement("div");
  div.style.padding = "6px";
  div.style.fontSize = "13px";
  div.style.borderLeft = "4px solid #0078d4";
  div.style.background = "#f1f5ff";
  div.style.marginBottom = "6px";
  div.textContent = text;
  box.appendChild(div);
  box.scrollTop = box.scrollHeight;
}

// Detect SP / Browser
const isSharePoint = typeof _spPageContextInfo !== "undefined";

// Init UI
document.addEventListener("DOMContentLoaded", () => {
  logMessage("✅ SmartCopilot ready");
  logMessage(`Mode: ${isSharePoint ? "SharePoint" : "Browser/GitHub"}`);

  const input = document.getElementById("searchBox");
  const send = document.getElementById("sendBtn");
  send.onclick = search;
  input.addEventListener("keypress", e => { if (e.key === "Enter") search(); });
});

async function search() {
  const input = document.getElementById("searchBox");
  const keyword = input.value.trim();

  if (!keyword) return;
  logMessage(`🔍 Search: ${keyword}`);

  if (!isSharePoint) {
    // Demo mode for GitHub
    logMessage("🧪 Demo mode");
    const demo = [
      {Title:"Số văn bản", Value:"{SoVB}", Desc:"Tự động số"},
      {Title:"Ngày ban hành", Value:"{NgayBH}", Desc:"Ngày ký"}
    ];
    showResults(demo, keyword);
    return;
  }

  try {
    const site = _spPageContextInfo.webAbsoluteUrl;
    const url = `${site}/_api/web/lists/getbytitle('TagLibrary')/items?$select=Title,Value,Desc`;

    logMessage("🌐 Call API");

    const r = await fetch(url, {
      headers: { Accept:"application/json;odata=verbose" },
      credentials:"same-origin"
    });

    if (!r.ok) throw new Error(`HTTP ${r.status}`);

    const json = await r.json();
    const rows = json.d.results;
    showResults(rows, keyword);

  } catch (e) {
    logMessage("❌ Lỗi: " + e.message);
  }
}

function showResults(items, keyword) {
  const match = items.filter(t =>
    t.Title.toLowerCase().includes(keyword.toLowerCase()) ||
    t.Value.toLowerCase().includes(keyword.toLowerCase()) ||
    (t.Desc || "").toLowerCase().includes(keyword.toLowerCase())
  );

  if (match.length === 0) {
    logMessage("⚠️ Không tìm thấy");
    return;
  }

  match.forEach(t => {
    logMessage(`📘 ${t.Title} → ${t.Value}`);
  });
}

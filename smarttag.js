// Detect environment (Office vs normal website)
const isOffice = typeof Office !== "undefined" && Office.onReady;

if (isOffice) {
  // Running inside Word Add-in
  Office.onReady(() => {
    initUI();
  });
} else {
  // Running on normal browser (GitHub testing)
  document.addEventListener("DOMContentLoaded", () => {
    console.log("✅ Running in browser mode (GitHub test)");
    initUI();
  });
}

// UI init
function initUI() {
  const input = document.getElementById("searchBox");
  const send = document.getElementById("sendBtn");

  if (!input || !send) {
    console.error("❌ UI not found – check HTML");
    return;
  }

  send.onclick = () => runSearch();
  input.addEventListener("keypress", e => { if (e.key === "Enter") runSearch(); });

  addMessage("✅ Smart Copilot sẵn sàng.\n• Word mode detected: " + isOffice + "\n• Gõ @ hoặc từ khóa để tra cứu TagLibrary", "system");
}

// User run search
async function runSearch() {
  const input = document.getElementById("searchBox");
  const keyword = input.value.trim();
  if (!keyword) return;

  addMessage(`🔎 Tìm kiếm: ${keyword}`, "system");
  await fetchTags(keyword);
  input.value = "";
}

// Fetch SharePoint list TagLibrary
async function fetchTags(keyword) {
  try {
    // Browser test mode → dummy values
    if (!isOffice && typeof _spPageContextInfo === "undefined") {
      addMessage("🌐 GitHub mode – demo data", "system");

      const demo = [
        { Title: "Số văn bản", Value: "{SoVB}", Desc: "Tự động điền số văn bản" },
        { Title: "Ngày ban hành", Value: "{NgayBanHanh}", Desc: "Ngày ký văn bản" }
      ];

      return showResults(keyword, demo);
    }

    // Real mode inside SharePoint
    const siteUrl = _spPageContextInfo.webAbsoluteUrl;
    const endpoint = `${siteUrl}/_api/web/lists/getbytitle('TagLibrary')/items?$select=Title,Value,Desc`;

    const response = await fetch(endpoint, {
      headers: { Accept: "application/json;odata=verbose" },
      credentials: "same-origin"
    });

    if (!response.ok) throw new Error(`HTTP ${response.status}`);

    const data = await response.json();
    const results = data.d?.results || [];

    showResults(keyword, results);

  } catch (err) {
    addMessage(`❌ Lỗi: ${err.message}`, "system");
  }
}

function showResults(keyword, results) {
  const filtered = results.filter(item =>
    !keyword ||
    item.Title.toLowerCase().includes(keyword.toLowerCase()) ||
    item.Value.toLowerCase().includes(keyword.toLowerCase()) ||
    (item.Desc || "").toLowerCase().includes(keyword.toLowerCase())
  );

  if (filtered.length === 0) {
    addMessage("⚠️ Không tìm thấy kết quả phù hợp.", "system");
    return;
  }

  filtered.forEach(tag => {
    addMessage(`📘 <b>${tag.Title}</b><br>${tag.Desc}<br><small>🔖 ${tag.Value}</small>`, "result");
  });
}

// Message UI helper
function addMessage(text, type = "system") {
  const container = document.getElementById("chat-output");
  const el = document.createElement("div");
  el.className = `message ${type}`;
  el.innerHTML = text.replace(/\n/g, "<br>");
  container.appendChild(el);
  container.scrollTop = container.scrollHeight;
}

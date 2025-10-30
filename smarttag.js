async function showDialog(message) {
  const dialog = document.getElementById("infoDialog");
  document.getElementById("dialogText").innerText = message;
  dialog.showModal();
}

async function searchTags(query) {
  const url = `https://justengineertech.sharepoint.com/sites/E-Office/_api/web/lists/getbytitle('TagLibrary')/items?$select=Title,Value,Desc&$top=50`;
  try {
    const response = await fetch(url, {
      headers: { "Accept": "application/json;odata=nometadata" }
    });

    if (!response.ok) {
      showDialog("Fetch lỗi HTTP: " + response.status);
      console.error("Fetch error:", response);
      return;
    }

    const data = await response.json();
    console.log("📦 Data SharePoint:", data);

    if (!data || !data.value || data.value.length === 0) {
      showDialog("Không tìm thấy dữ liệu");
      return;
    }

    const filtered = data.value.filter(
      item => item.Title.toLowerCase().includes(query.toLowerCase())
    );

    const list = document.getElementById("resultList");
    list.innerHTML = "";
    if (filtered.length > 0) {
      filtered.forEach(item => {
        const li = document.createElement("li");
        li.textContent = `${item.Title} (${item.Value})`;
        li.title = item.Desc;
        list.appendChild(li);
      });
    } else {
      list.innerHTML = "<li>Không tìm thấy kết quả</li>";
    }
  } catch (err) {
    console.error("❌ Lỗi JS:", err);
    showDialog("Lỗi JavaScript hoặc CORS.");
  }
}

document.getElementById("searchBox").addEventListener("input", (e) => {
  const query = e.target.value.trim();
  if (query.length > 1) searchTags(query);
});

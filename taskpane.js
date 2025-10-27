let tags = [];

Office.onReady(() => {
  console.log("Office.js ready");
  loadTags();
});

async function loadTags() {
  const loading = document.getElementById("loading");
  loading.textContent = "Đang tải dữ liệu tag từ SharePoint...";

  try {
    const response = await fetch(
      "https://justengineertech.sharepoint.com/sites/E-Office/_api/web/lists/getbytitle('TagLibrary')/items?$select=Title,Value,Description",
      {
        headers: {
          Accept: "application/json;odata=verbose"
        },
        credentials: "include"
      }
    );

    const data = await response.json();
    tags = data.d.results.map(item => ({
      name: item.Title,
      value: item.Value,
      description: item.Description
    }));

    loading.textContent = `✅ Đã tải ${tags.length} tag`;
  } catch (err) {
    console.error("Lỗi tải tag:", err);
    loading.textContent = "❌ Không tải được dữ liệu từ SharePoint";
  }
}

// Xử lý nhập @tag
document.getElementById("tagSearch").addEventListener("input", function (e) {
  const keyword = e.target.value.toLowerCase();
  const suggestionBox = document.getElementById("suggestions");
  suggestionBox.innerHTML = "";

  if (keyword.startsWith("@")) {
    const results = tags.filter(t => t.name.toLowerCase().includes(keyword));
    results.forEach(tag => {
      const div = document.createElement("div");
      div.className = "suggestion-item";
      div.innerHTML = `<strong>${tag.name}</strong> - ${tag.value}<br><small>${tag.description}</small>`;
      div.onclick = () => insertTag(tag.value);
      suggestionBox.appendChild(div);
    });
  }
});

async function insertTag(value) {
  try {
    await Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.insertText(value, "Replace");
      await context.sync();
    });
  } catch (error) {
    console.error("Insert error:", error);
    alert("Không thể chèn tag vào Word. Vui lòng thử lại.");
  }
}

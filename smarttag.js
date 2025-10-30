const siteUrl = "https://justengineertech.sharepoint.com/sites/E-Office";
const listName = "TagLibrary";

document.getElementById("searchBox").addEventListener("input", async (e) => {
  const keyword = e.target.value.trim().toLowerCase();
  const resultsList = document.getElementById("tagResults");
  resultsList.innerHTML = "";
  if (!keyword) return;

  try {
    const response = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$select=Title,Value,Desc`, {
      headers: { "Accept": "application/json;odata=verbose" }
    });
    const data = await response.json();
    const items = data.d?.results || [];

    // Lọc tương đối
    const filtered = items.filter(i => i.Title.toLowerCase().includes(keyword));

    if (filtered.length === 0) {
      // Nếu không có kết quả tương đối → kiểm tra trùng tuyệt đối
      const exact = items.find(i => i.Title.toLowerCase() === keyword);
      resultsList.innerHTML = exact
        ? `<li title="${exact.Value}\n${exact.Desc}">${exact.Title}</li>`
        : "<li>Không tìm thấy tag</li>";
      return;
    }

    // Hiển thị kết quả tương đối
    filtered.forEach(i => {
      const li = document.createElement("li");
      li.textContent = i.Title;
      li.title = `${i.Value}\n${i.Desc}`;
      resultsList.appendChild(li);
    });

    console.log("Kết quả tìm kiếm:", filtered);
  } catch (error) {
    console.error("Lỗi khi tải tag:", error);
  }
});

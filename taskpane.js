Office.onReady(() => {
  document.getElementById("searchBox").addEventListener("input", searchTag);
});

async function searchTag(e) {
  const query = e.target.value.trim();
  if (!query) {
    document.getElementById("results").innerHTML = "";
    return;
  }

  const listName = "TagLibrary";
  const siteUrl = "https://justengineertech.sharepoint.com/sites/E-Office";

  try {
    const response = await fetch(`${siteUrl}/_api/web/lists/getbytitle('${listName}')/items?$filter=startswith(Title,'${query}')`, {
      headers: { "Accept": "application/json;odata=verbose" }
    });
    const data = await response.json();
    const results = data.d.results;

    let html = "";
    for (let item of results) {
      html += `<div onclick="insertTag('${item.Title}','${item.Value}','${item.Description}')">
                 <h3>${item.Title}</h3>
                 <p><b>Giá trị:</b> ${item.Value || "-"}<br><b>Mô tả:</b> ${item.Description || "-"}</p>
               </div>`;
    }
    document.getElementById("results").innerHTML = html || "<p>Không tìm thấy thẻ nào.</p>";

  } catch (err) {
    console.error(err);
    document.getElementById("results").innerHTML = "<p>Lỗi khi truy vấn SharePoint.</p>";
  }
}

function insertTag(name, value, desc) {
  Office.context.document.setSelectedDataAsync(`${value}`, () => {
    console.log("Đã chèn:", value);
  });
}

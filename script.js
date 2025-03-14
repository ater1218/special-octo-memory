const SCRIPT_URL = "https://script.google.com/macros/s/YOUR_SCRIPT_ID/exec"; // Google Apps ScriptのウェブアプリURLに置き換え
let headers = [];

function showLoading(show) {
  document.getElementById("loading").style.display = show ? "block" : "none";
}

function fetchSheetNames() {
  showLoading(true);
  const url = document.getElementById("spreadsheetUrl").value;
  const spreadsheetId = url.match(/\/d\/(.+?)\//)[1];

  fetch(`${SCRIPT_URL}?spreadsheetId=${spreadsheetId}&action=getSheetNames`)
    .then(response => response.json())
    .then(result => {
      const select = document.getElementById("sheetName");
      select.innerHTML = '<option value="">-- シートを選択 --</option>';
      result.sheetNames.forEach(name => {
        const option = document.createElement("option");
        option.value = name;
        option.text = name;
        select.appendChild(option);
      });
    })
    .catch(error => console.error("シート名取得エラー:", error))
    .finally(() => showLoading(false));
}

function fetchData() {
  const url = document.getElementById("spreadsheetUrl").value;
  const sheetName = document.getElementById("sheetName").value;
  if (!sheetName) return;

  showLoading(true);
  const spreadsheetId = url.match(/\/d\/(.+?)\//)[1];

  fetch(`${SCRIPT_URL}?spreadsheetId=${spreadsheetId}&sheetName=${sheetName}`)
    .then(response => response.json())
    .then(result => {
      if (result.headers && Array.isArray(result.headers)) {
        headers = result.headers;
        document.getElementById("headers").innerText = "ヘッダー: " + headers.join(", ");
      } else {
        console.error("ヘッダーが取得できませんでした");
      }

      if (result.bValues && Array.isArray(result.bValues)) {
        const select = document.getElementById("keyValue");
        select.innerHTML = "";
        result.bValues.forEach(value => {
          const option = document.createElement("option");
          option.value = value;
          option.text = value;
          select.appendChild(option);
        });
      } else {
        console.error("B列の値が取得できませんでした");
      }

      if (headers && headers.length > 0) {
        let inputs = "";
        headers.forEach(header => {
          inputs += `<label>${header}: <input type="text" id="input-${header}"></label><br>`;
        });
        document.getElementById("inputFields").innerHTML = inputs;
      }
    })
    .catch(error => console.error("データ取得エラー:", error))
    .finally(() => showLoading(false));
}

function submitData() {
  showLoading(true);
  const url = document.getElementById("spreadsheetUrl").value;
  const sheetName = document.getElementById("sheetName").value;
  const spreadsheetId = url.match(/\/d\/(.+?)\//)[1];
  const keyValue = document.getElementById("keyValue").value;

  const data = {};
  headers.forEach(header => {
    const value = document.getElementById(`input-${header}`).value;
    if (value) data[header] = value;
  });

  fetch(SCRIPT_URL, {
    method: "POST",
    body: JSON.stringify({
      spreadsheetId: spreadsheetId,
      sheetName: sheetName,
      key: keyValue,
      data: data
    })
  })
  .then(response => response.text())
  .then(result => alert(result))
  .catch(error => console.error("送信エラー:", error))
  .finally(() => showLoading(false));
}

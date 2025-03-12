const SCRIPT_URL = "https://script.google.com/macros/s/AKfycbx1pOAPlQFcCwG6paZylu88JU0leDVaqRgfk5VtTgrOKTnBVEWJ9hgFRHcSDkK4seLTlA/exec"; // Google Apps ScriptのウェブアプリURLに置き換え
let headers = [];

function fetchSheetNames() {
  const url = document.getElementById("spreadsheetUrl").value;
  const spreadsheetId = url.match(/\/d\/(.+?)\//)[1];

  fetch(`${SCRIPT_URL}?spreadsheetId=${spreadsheetId}&action=getSheetNames`)
    .then(response => response.json())
    .then(result => {
      const select = document.getElementById("sheetName");
      select.innerHTML = "";
      result.sheetNames.forEach(name => {
        const option = document.createElement("option");
        option.value = name;
        option.text = name;
        select.appendChild(option);
      });
    })
    .catch(error => console.error("シート名取得エラー:", error));
}

function fetchData() {
  const url = document.getElementById("spreadsheetUrl").value;
  const sheetName = document.getElementById("sheetName").value;
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
    .catch(error => console.error("データ取得エラー:", error));
}

function submitData() {
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
  .catch(error => console.error("送信エラー:", error));
}

function setupCallLink() {
  const num = document.getElementById("userInput").value.trim();
  const container = document.getElementById("callLinkArea");
  container.innerHTML = "";

  const a = document.createElement("a");
  a.href = num ? "tel:0120945779,,,,," + num + "#" : "#";
  a.innerText = "📞 代理店コールセンターへ発信";
  a.style = "display:inline-block; padding: 6px 14px; background-color: #007cba; color: white; border: none; border-radius: 4px; text-decoration: none; font-size: 15px;";
  a.onclick = function() {
    if (!num) {
      alert("募集人番号を入力してください");
      return false;
    }
  };
  container.appendChild(a);
}

function showCallInterface() {
  document.getElementById("formBlock").style.display = "block";
  document.getElementById("linkContainer").style.display = "block";
  setupCallLink();
}

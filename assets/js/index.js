const inputElement = document.querySelector(".email");
const formSubmit = document.querySelector(".submitBtn");
const tbody = document.getElementById("tbody");
const exportData = document.getElementById("export");

window.onload = loadData();

formSubmit.addEventListener("submit", (e) => {
  e.preventDefault();
  const time = new Date().toDateString();
  const value = inputElement.value;
  if ((value && value.trim()) !== "") {
    const leadsSaved = localStorage.getItem("leadsSaved")
      ? JSON.parse(localStorage.getItem("leadsSaved"))
      : [];
    const leadsData = {
      leadsEmail: value,
      date: time,
    };
    leadsSaved.unshift(leadsData);
    localStorage.setItem("leadsSaved", JSON.stringify(leadsSaved));
    tbody.innerHTML = "";
    inputElement.value = "";
    loadData();
  }
});

function loadData() {
  const leadsSaved = localStorage.getItem("leadsSaved")
    ? JSON.parse(localStorage.getItem("leadsSaved"))
    : [];
  if (leadsSaved.length !== 0) {
    leadsSaved.forEach((Element) => displayValues(Element.leadsEmail));
  }
}

function displayValues(email) {
  const tr = document.createElement("tr");
  const td1 = document.createElement("td");
  td1.textContent = email;
  tr.appendChild(td1);
  tbody.appendChild(tr);
}

exportData.addEventListener("click", (e) => {
  e.preventDefault();
  const leadsSaved =
    localStorage.getItem("leadsSaved") &&
    JSON.parse(localStorage.getItem("leadsSaved"));
  if (!!leadsSaved) {
    const worksheet = XLSX.utils.json_to_sheet(leadsSaved);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Dates");

    /* fix headers */
    XLSX.utils.sheet_add_aoa(worksheet, [["Email of Lead"]], { origin: "A1" });

    /* calculate column width */
    const max_width = leadsSaved.reduce((w, r) => Math.max(w, r.leadsEmail.length), 10);
    worksheet["!cols"] = [ { wch: max_width } ];

    /* create an XLSX file and try to save to Presidents.xlsx */
    XLSX.writeFile(workbook, "Leads.xlsx", { compression: true });
  }
});

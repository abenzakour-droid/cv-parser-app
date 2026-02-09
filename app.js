const fileInput = document.getElementById("file-input");
const dropZone = document.getElementById("drop-zone");
const statusText = document.getElementById("status-text");
const fileNameTag = document.getElementById("file-name");
const fileTypeTag = document.getElementById("file-type");
const scanScoreTag = document.getElementById("scan-score");
const resultsBody = document.getElementById("results-body");
const exportButton = document.getElementById("export-button");

const CONTACT_FIELDS = ["Nom complet", "Email", "Téléphone", "Ville / Pays", "LinkedIn"];

const updateStatus = (message, detail = {}) => {
  statusText.textContent = message;
  if (detail.fileName) fileNameTag.textContent = detail.fileName;
  if (detail.fileType) fileTypeTag.textContent = detail.fileType;
  if (detail.scanScore) scanScoreTag.textContent = detail.scanScore;
};

const setTableRow = (values) => {
  resultsBody.innerHTML = "";
  const row = document.createElement("tr");

  values.forEach((value) => {
    const cell = document.createElement("td");
    const input = document.createElement("input");
    input.value = value || "";
    input.placeholder = "—";
    cell.appendChild(input);
    row.appendChild(cell);
  });

  resultsBody.appendChild(row);
};

const extractContact = (text) => {
  const normalized = text.replace(/\s+/g, " ").trim();
  const email = normalized.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i)?.[0] || "";
  const phone = normalized.match(/(\+\d{1,3}[\s-]?)?(\(?\d{2,4}\)?[\s-]?){2,4}\d{2,4}/)?.[0] || "";
  const linkedin = normalized.match(/(https?:\/\/)?(www\.)?linkedin\.com\/[A-Za-z0-9_\-\/]+/i)?.[0] || "";

  const firstLine = text.split(/\n/).find((line) => line.trim().length > 2) || "";
  const locationMatch = normalized.match(/(Casablanca|Rabat|Marrakesh|Fes|Tanger|Agadir|Morocco|Maroc|France|Paris|Lyon|Marseille)/i);

  return {
    name: firstLine.trim().slice(0, 40),
    email,
    phone,
    location: locationMatch ? locationMatch[0] : "",
    linkedin,
  };
};

const readPdfText = async (file) => {
  const arrayBuffer = await file.arrayBuffer();
  const pdf = await window.pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  let text = "";
  for (let i = 1; i <= pdf.numPages; i += 1) {
    const page = await pdf.getPage(i);
    const content = await page.getTextContent();
    text += content.items.map((item) => item.str).join(" ") + "\n";
  }
  return text;
};

const readDocxText = async (file) => {
  const arrayBuffer = await file.arrayBuffer();
  const result = await window.mammoth.extractRawText({ arrayBuffer });
  return result.value;
};

const handleFile = async (file) => {
  if (!file) return;
  const extension = file.name.split(".").pop()?.toLowerCase();

  updateStatus("Analyse en cours...", {
    fileName: file.name,
    fileType: extension === "pdf" ? "PDF" : "DOCX",
    scanScore: "Analyse",
  });

  try {
    let text = "";
    if (extension === "pdf") {
      text = await readPdfText(file);
    } else if (extension === "docx") {
      text = await readDocxText(file);
    } else {
      updateStatus("Format non supporté.", { scanScore: "Erreur" });
      return;
    }

    const contact = extractContact(text);
    const values = [contact.name, contact.email, contact.phone, contact.location, contact.linkedin];
    setTableRow(values);
    exportButton.disabled = false;
    updateStatus("Analyse terminée. Vérifiez les champs ci-dessous.", {
      scanScore: "OK",
    });
  } catch (error) {
    console.error(error);
    updateStatus("Impossible d'analyser le fichier. Réessayez.", {
      scanScore: "Erreur",
    });
  }
};

const exportToExcel = () => {
  const row = resultsBody.querySelector("tr");
  if (!row) return;

  const values = Array.from(row.querySelectorAll("input")).map((input) => input.value);
  const worksheet = XLSX.utils.aoa_to_sheet([CONTACT_FIELDS, values]);
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, "Coordonnées");
  XLSX.writeFile(workbook, "cv-coordonnees.xlsx");
};

fileInput.addEventListener("change", (event) => handleFile(event.target.files[0]));
exportButton.addEventListener("click", exportToExcel);

["dragenter", "dragover"].forEach((eventName) => {
  dropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropZone.classList.add("active");
  });
});

["dragleave", "drop"].forEach((eventName) => {
  dropZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    dropZone.classList.remove("active");
  });
});

dropZone.addEventListener("drop", (event) => {
  const [file] = event.dataTransfer.files;
  handleFile(file);
});

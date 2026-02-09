const fileInput = document.getElementById("file-input");
const dropZone = document.getElementById("drop-zone");
const statusText = document.getElementById("status-text");
const fileNameTag = document.getElementById("file-name");
const fileTypeTag = document.getElementById("file-type");
const scanScoreTag = document.getElementById("scan-score");
const progressBar = document.getElementById("progress-bar");
const resultsBody = document.getElementById("results-body");
const exportButton = document.getElementById("export-button");
const copyButton = document.getElementById("copy-button");
const chooseFileButton = document.getElementById("choose-file-button");
const resetButton = document.getElementById("reset-button");

const CONTACT_FIELDS = ["Nom complet", "Email", "Téléphone", "Ville / Pays", "LinkedIn"];

const updateStatus = (message, detail = {}) => {
  statusText.textContent = message;
  if (detail.fileName) fileNameTag.textContent = detail.fileName;
  if (detail.fileType) fileTypeTag.textContent = detail.fileType;
  if (detail.scanScore) scanScoreTag.textContent = detail.scanScore;
  if (typeof detail.progress === "number") {
    progressBar.style.width = `${detail.progress}%`;
  }
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

const resetTable = () => {
  setTableRow([
    "Ex: Salma Ait Lahcen",
    "exemple@mail.com",
    "+212 6 12 34 56 78",
    "Casablanca, Maroc",
    "linkedin.com/in/salma",
  ]);
  exportButton.disabled = true;
  copyButton.disabled = true;
  resetButton.disabled = true;
  updateStatus("Aucun fichier importé. Aperçu affiché.", {
    fileName: "—",
    fileType: "—",
    scanScore: "Aperçu",
    progress: 0,
  });
};

const extractContact = (text) => {
  const normalized = text.replace(/\s+/g, " ").trim();
  const email = normalized.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i)?.[0] || "";
  const phone =
    normalized.match(/(\+\d{1,3}[\s-]?)?(\(?\d{2,4}\)?[\s-]?){2,4}\d{2,4}/)?.[0] || "";
  const linkedin = normalized.match(/(https?:\/\/)?(www\.)?linkedin\.com\/[A-Za-z0-9_\-\/]+/i)?.[0] || "";

  const lines = text
    .split(/\n/)
    .map((line) => line.trim())
    .filter((line) => line.length > 2);
  const firstLine =
    lines.find((line) => !line.toLowerCase().includes("curriculum") && !line.includes("@")) || "";
  const locationMatch = normalized.match(
    /(Casablanca|Rabat|Marrakesh|Fes|Tanger|Agadir|Morocco|Maroc|France|Paris|Lyon|Marseille)/i
  );

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
    progress: 20,
  });

  try {
    let text = "";
    if (extension === "pdf") {
      updateStatus("Lecture du PDF...", { scanScore: "Analyse", progress: 40 });
      text = await readPdfText(file);
    } else if (extension === "docx") {
      updateStatus("Lecture du document Word...", { scanScore: "Analyse", progress: 40 });
      text = await readDocxText(file);
    } else {
      updateStatus("Format non supporté.", { scanScore: "Erreur", progress: 0 });
      return;
    }

    updateStatus("Extraction des coordonnées...", { scanScore: "Analyse", progress: 70 });
    const contact = extractContact(text);
    const values = [contact.name, contact.email, contact.phone, contact.location, contact.linkedin];
    setTableRow(values);
    exportButton.disabled = false;
    copyButton.disabled = false;
    resetButton.disabled = false;
    updateStatus("Analyse terminée. Vérifiez les champs ci-dessous.", {
      scanScore: "OK",
      progress: 100,
    });
  } catch (error) {
    console.error(error);
    updateStatus("Impossible d'analyser le fichier. Réessayez.", {
      scanScore: "Erreur",
      progress: 0,
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

const copyToClipboard = async () => {
  const row = resultsBody.querySelector("tr");
  if (!row) return;
  const values = Array.from(row.querySelectorAll("input")).map((input) => input.value);
  const data = CONTACT_FIELDS.map((field, index) => `${field}: ${values[index] || "—"}`).join("\n");
  try {
    await navigator.clipboard.writeText(data);
    updateStatus("Coordonnées copiées dans le presse-papiers.", {
      scanScore: "OK",
    });
  } catch (error) {
    console.error(error);
    updateStatus("Impossible de copier. Essayez manuellement.", {
      scanScore: "Erreur",
    });
  }
};

fileInput.addEventListener("change", (event) => handleFile(event.target.files[0]));
exportButton.addEventListener("click", exportToExcel);
copyButton.addEventListener("click", copyToClipboard);
chooseFileButton.addEventListener("click", () => fileInput.click());
resetButton.addEventListener("click", resetTable);

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

resetTable();

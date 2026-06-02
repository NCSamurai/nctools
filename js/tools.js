function generateList(bin, wo) {
  const piTable = document.getElementById("piTable");
  const tr = document.createElement("tr");
  tr.setAttribute("id", wo);

  const binContent = document.createElement("td");
  binContent.innerText = bin;

  const woContent = document.createElement("td");
  woContent.className = "wo";
  woContent.innerText = wo;

  tr.appendChild(binContent);
  tr.appendChild(woContent);
  piTable.appendChild(tr);
}
function generateBarCode(val) {
  document.getElementById("showBarcode").style.display = "block";

  JsBarcode("#barcode", val, {
    format: "CODE128",
    lineColor: "#000",
    width: 4,
    height: 60,
    displayValue: true,
  });
}
function waitForClick(next, previous) {
  return new Promise((resolve) => {
    const nextHandler = () => {
      cleanup();
      resolve("next");
    };
    
    const prevHandler = () => { 
      cleanup();
      resolve("previous");
    };

    function cleanup() {
      next.removeEventListener("click", nextHandler);
      previous.removeEventListener("click", prevHandler);
    }

    next.addEventListener("click", nextHandler);
    previous.addEventListener("click", prevHandler);
  });
}
async function woAutomation(warehouseOrders) {
  const next = document.getElementById("next");
  const prev = document.getElementById("previous");

  let index = 0;

  while (index >= 0 && index < warehouseOrders.length) {
    generateBarCode(warehouseOrders[index]);
    const currentBin = document.getElementById(warehouseOrders[index]);
    currentBin.style.backgroundColor = "#5b0000";
    const action = await waitForClick(next, prev);
    
    if (action === "next") {
      currentBin.style.backgroundColor = "#004e03";
      index++;
    }
    if (action === "previous") {
      currentBin.style.backgroundColor = "inherit";
      index--;
    }
  }
  document.getElementById("showBarcode").style.display = "none";
  document.getElementById("sectionComplete").style.display = "block";
  
  const done = document.getElementById("taskComplete");
  const cancel = document.getElementById("taskCancel");

  const confirmation = await waitForClick(done, cancel);

  if (confirmation === "next") {
    document.getElementById("sectionComplete").style.display = "none";
    document.getElementById(currentSection).style.display = "none";
  }
  if (confirmation === "previous")
    document.getElementById("sectionComplete").style.display = "none";
}

const pdfInput = document.getElementById("pdfInput");
const excelInput = document.getElementById("excelInput");
const downloads = document.getElementById("downloads");
const downloadBtn = document.getElementById("downloadBtn");
const piTable = document.getElementById("piTable");
const start = document.getElementById("start-auto");
const stop = document.getElementById("stop-auto");
const tasks = new Map();
let output;
let currentSection;

pdfInput.addEventListener("change", async (e) => {
  const files = Array.from(e.target.files);
  downloads.innerHTML = "";

  for (const file of files) {
    const buffer = await file.arrayBuffer();
    const sourcePDF = await PDFLib.PDFDocument.load(buffer, {
      ignoreEncryption: true,
    });
    const unlockedPDF = await PDFLib.PDFDocument.create();

    const pages = await unlockedPDF.copyPages(
      sourcePDF,
      sourcePDF.getPageIndices(),
    );
    pages.forEach((page) => unlockedPDF.addPage(page));

    output = await unlockedPDF.save();

    const blob = new Blob([output], { type: "application/pdf" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");

    a.href = url;
    a.className = "links";
    a.download = `u${file.name}`;
    a.textContent = `Download ${file.name}`;
    a.style.display = "block";
    downloads.appendChild(a);
    // a.click(); // to automatically download - test to not show dialog box
  }
});
excelInput.addEventListener("change", async (e) => {
  const file = e.target.files[0];
  if (!file) return;

  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });
  const sheetName = workbook.SheetNames[0];
  const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
  let sectionList = new Set();

  for (let row = 0; row < data.length; row++) {
    if (data[row]["Storage Bin"] === "") break;

    const bin = data[row]["Storage Bin"];
    const order = data[row]["Warehouse Order"];
    const section = bin.slice(0, 3);

    sectionList.add(section);
    tasks.set(bin, order);
  }
  for (let section of sectionList) {
    const content = document.getElementById("section-content");
    const button = document.createElement("button");

    button.innerText = section;
    button.className = "button sectionBtns";
    button.setAttribute("id", section);
    content.appendChild(button);
  }
});
document.addEventListener("click", (e) => {
  if (e.target.classList.contains("sectionBtns")) {
    const section = e.target.innerText;
    currentSection = section;
    const taskSections = tasks.keys();

    piTable.innerHTML = "";

    const thBin = document.createElement("th");
    thBin.innerText = "Bin";
    const thWo = document.createElement("th");
    thWo.innerText = "Warehouse Order";
    const thBarcode = document.createElement("th");
    thBarcode.innerText = "Barcode";

    piTable.appendChild(thBin);
    piTable.appendChild(thWo);

    for (let key of taskSections) {
      const bin = key.slice(0, 3);
      if (bin === section) {
        generateList(key, tasks.get(key));
      }
    }
  }
});
start.addEventListener("click", () => {
  const warehouseOrders = [...document.querySelectorAll(".wo")].map(
    (el) => el.innerHTML,
  );
  woAutomation(warehouseOrders);
});
stop.addEventListener("click", () => { 
  document.getElementById("showBarcode").style.display = "none";
});
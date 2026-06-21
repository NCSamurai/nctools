function generateList(bin, wo, pi) {
  const piTable = document.getElementById("piTable");
  const tr = document.createElement("tr");
  tr.setAttribute("id", wo);

  const binContent = document.createElement("td");
  binContent.className = "bin";
  binContent.innerText = bin;

  const woContent = document.createElement("td");
  woContent.className = "wo";
  woContent.innerText = wo;

  const piContent = document.createElement("td");
  piContent.innerText = pi;

  tr.appendChild(binContent);
  tr.appendChild(woContent);
  tr.appendChild(piContent);

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
  const copy = document.getElementById("copy-wo");
  const close = document.getElementById("cancelX-btn");

  return new Promise((resolve) => {
    const nextHandler = () => {
      cleanup();
      resolve("next");
    };
    const prevHandler = () => {
      cleanup();
      resolve("previous");
    };
    const copyHandler = () => {
      cleanup();
      resolve("copy");
    };
    const closeHandler = () => {
      cleanup();
      resolve("close");
    };

    function cleanup() {
      next.removeEventListener("click", nextHandler);
      previous.removeEventListener("click", prevHandler);
      copy.removeEventListener("click", copyHandler);
      close.removeEventListener("click", closeHandler);
    }

    next.addEventListener("click", nextHandler);
    previous.addEventListener("click", prevHandler);
    copy.addEventListener("click", copyHandler);
    close.addEventListener("click", closeHandler);
  });
}
async function woAutomation(warehouseOrders) {
  const next = document.getElementById("next");
  const prev = document.getElementById("previous");
  const copy = document.getElementById("copy-wo");
  const close = document.getElementById("cancelX-btn");

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
    if (action === "copy") {
      navigator.clipboard.writeText(currentBin.getAttribute("id"));
    }
    if (action === "close") {
      document.getElementById("showBarcode").style.display = "none";
      return;
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
function copyAllBins() {
  const values = [...document.querySelectorAll(".bin")].map(
    (el) => el.innerHTML,
  );
  const bins = values.join("\n");

  navigator.clipboard.writeText(bins);
}

const excelInput = document.getElementById("excelInput");
const piTable = document.getElementById("piTable");
const start = document.getElementById("start-auto");
const copyBins = document.getElementById("copy-bins");
const order = document.getElementById("changeOrder");
const tasks = new Map();
let output;
let currentSection;

excelInput.addEventListener("change", async (e) => {
  const file = e.target.files[0];
  if (!file) return;

  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });
  const sheetName = workbook.SheetNames[0];
  const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
  data.sort((a, b) => {
    if (a["Storage Bin"] < b["Storage Bin"]) return -1;
    if (a["Storage Bin"] > b["Storage Bin"]) return 1;
    return 0;
  });
  let sectionList = new Set();

  for (let row = 0; row < data.length; row++) {
    if (data[row]["Storage Bin"] === undefined) break;
      
    const bin = data[row]["Storage Bin"];
    const order = data[row]["Warehouse Order"];
    const pi = data[row]["Physical Inventory Document"];
    const section = bin.slice(0, 3);

    sectionList.add(section);
    tasks.set(bin, [order, pi]);
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
    const thPi = document.createElement("th");
    thPi.innerText = "PI Document";

    piTable.appendChild(thBin);
    piTable.appendChild(thWo);
    piTable.appendChild(thPi);

    for (let key of taskSections) {
      const bin = key.slice(0, 3);
      const val = tasks.get(key);
      if (bin === section) {
        generateList(key, val[0], val[1]);
      }
    }
  }
});


// Action Bar Functions
start.addEventListener("click", () => {
  const warehouseOrders = [...document.querySelectorAll(".wo")].map(
    (el) => el.innerHTML,
  );
  woAutomation(warehouseOrders);
});
copyBins.addEventListener("click", () => {
  copyBins.classList.remove("active");
  void copyBins.offsetWidth;
  copyBins.classList.add("active");
  copyAllBins();
});
order.addEventListener("click", () => {
  const values = [...document.querySelectorAll(".bin")]
    .map((el) => el.innerHTML).reverse();

  piTable.innerHTML = "";

  const thBin = document.createElement("th");
  thBin.innerText = "Bin";
  const thWo = document.createElement("th");
  thWo.innerText = "Warehouse Order";
  const thPi = document.createElement("th");
  thPi.innerText = "PI Document";

  piTable.appendChild(thBin);
  piTable.appendChild(thWo);
  piTable.appendChild(thPi);

  for (let key of values) {
    const bin = key.slice(0, 3);
    const val = tasks.get(key);
    if (bin === currentSection) {
      generateList(key, val[0], val[1]);
    }
  }
});
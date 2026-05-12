function toEWM(val) {
  const ewmStart = "10044710";
  const digits = String(val).slice(0, 9);
  const weights = [3, 1, 3, 1, 3, 1, 3, 1, 3];

  const total =
    29 +
    digits.split("").reduce((sum, d, i) => sum + parseInt(d) * weights[i], 0);
  const checkDigit = (10 - (total % 10)) % 10;

  return `${ewmStart}${digits + String(checkDigit)}`;
}

function convertTicket() {
  const ticketInput = document
    .getElementById("ticket-num")
    .value.trim()
    .split("\n");
  const multiInput = document
    .getElementById("multi-ticket")
    .value.trim()
    .split("\n");

  const inputs = [...new Set([...ticketInput, ...multiInput])].filter(Boolean);

  for (let ticket of inputs) {
    ticket = ticket.trim();
    const notValid = isNaN(parseInt(ticket));

    const isValid =
      ticket.length == 10 || ticket.length == 18 || ticket.length == 20;

    if (!isValid || notValid) {
      alert("Enter a valid ticket");
      return;
    }

    let ewmTicket;
    if (ticket.length == 18) ewmTicket = ticket;
    else if (ticket.length == 20) ewmTicket = ticket.slice(2);
    else ewmTicket = toEWM(ticket);

    const matricsTicket = `${ticket.slice(-10, -1)}0`;

    addTicket(matricsTicket, ewmTicket);
  }
}

function addTicket(matrics, ewm) {
  const table = document.getElementById("ticketTable");
  const tr = document.createElement("tr");

  const codeMatrics = document.createElement("td");
  codeMatrics.innerHTML = `<button class="button button-confirm barcodeBtns" data-value="${matrics}">Show</button>`;

  const codeEwm = document.createElement("td");
  codeEwm.innerHTML = `<button class="button button-confirm barcodeBtns" data-value="${ewm}">Show</button>`;

  const tdMatrics = document.createElement("td");
  tdMatrics.setAttribute("class", "matrics");
  tdMatrics.textContent = matrics;

  const tdEwm = document.createElement("td");
  tdEwm.setAttribute("class", "ewm");
  tdEwm.textContent = ewm;

  tr.appendChild(codeMatrics);
  tr.appendChild(tdMatrics);
  tr.appendChild(tdEwm);
  tr.appendChild(codeEwm);

  table.appendChild(tr);
}

function copyMatrics() {
  const values = [...document.querySelectorAll(".matrics")].map(
    (el) => el.innerHTML,
  );
  const tickets = values.join("\n");

  navigator.clipboard.writeText(tickets);
}

function copyEWM() {
  const values = [...document.querySelectorAll(".ewm")].map(
    (el) => el.innerHTML,
  );
  const tickets = values.join("\n");

  navigator.clipboard.writeText(tickets);
}

function clearTable() {
  const table = document.getElementById("ticketTable");
  table.innerHTML = "";

  const tr = document.createElement("tr");

  const codeMatrics = document.createElement("th");
  codeMatrics.textContent = "Barcode";

  const thMatrics = document.createElement("th");
  thMatrics.textContent = "Matrics";

  const thEwm = document.createElement("th");
  thEwm.textContent = "EWM";

  const codeEwm = document.createElement("th");
  codeEwm.textContent = "Barcode";

  tr.appendChild(codeMatrics);
  tr.appendChild(thMatrics);
  tr.appendChild(thEwm);
  tr.appendChild(codeEwm);

  table.appendChild(tr);
}

function startBarcodeScan(targetInputId) {
  const scannerModal = document.getElementById("scannerModal");
  const scannerEl = document.getElementById("scanner");
  let scanLocked = false;

  scannerModal.style.display = "block";
  Quagga.offDetected();

  Quagga.init(
    {
      inputStream: {
        type: "LiveStream",
        target: scannerEl,
        constraints: {
          facingMode: "environment",
        },
      },
      decoder: {
        readers: ["code_128_reader"],
      },
      locate: true,
    },
    (err) => {
      if (err) {
        console.error("Quagga init error:", err);
        return;
      }
      Quagga.start();
    },
  );

  Quagga.onDetected((result) => {
    if (scanLocked) return;

    const code = result?.codeResult?.code;
    if (!code || !/^\d{6,10}$/.test(code)) return; // only accept 6–10 digit codes

    scanLocked = true;
    document.getElementById(targetInputId).value = code;

    Quagga.stop();
    Quagga.offDetected();
    scannerModal.style.display = "none";
  });
}

function createBarCode(val) {
  document.getElementById("showBarcode").style.display = "block";

  JsBarcode("#barcode", val, {
    format: "CODE128",
    lineColor: "#000",
    width: 4,
    height: 60,
    displayValue: true,
  });
}

document.getElementById("multi-ticket-btn").addEventListener("click", () => {
  let button = document.getElementById("multi-ticket-btn");

  const currentState = document
    .getElementById("multi-ticket")
    .getAttribute("style");

  const isOn = currentState === "display: block;";

  if (isOn) {
    document
      .getElementById("multi-ticket")
      .setAttribute("style", "display: none;");
    document.getElementById("multi-ticket").value = "";
    document.getElementById("ticket-num").removeAttribute("disabled");
    button.innerText = "Add Multiple";
  } else {
    document
      .getElementById("multi-ticket")
      .setAttribute("style", "display: block;");
    document.getElementById("ticket-num").value = "";
    document.getElementById("ticket-num").setAttribute("disabled", true);
    button.innerText = "Add Single";
  }
});
document.getElementById("convert-ticket-btn").addEventListener("click", () => {
  convertTicket();
});
document.getElementById("hideBarcode").addEventListener("click", () => {
  document.getElementById("showBarcode").style.display = "none";
});
document.addEventListener("click", (e) => {
  if (e.target.classList.contains("barcodeBtns")) {
    e.target.classList.remove("active");
    void e.offsetWidth;
    e.target.classList.add("active");

    const ticket = e.target.dataset.value;
    createBarCode(ticket);
  }
});
document.getElementById("clearTable").addEventListener("click", () => {
  clearTable();
});
document.getElementById("copyMatricsBtn").addEventListener("click", () => {
  const button = document.getElementById("copyMatricsBtn");
  button.classList.remove("active");
  void button.offsetWidth;
  button.classList.add("active");
  copyMatrics();
});
document.getElementById("copyEWMBtn").addEventListener("click", () => {
  const button = document.getElementById("copyEWMBtn");
  button.classList.remove("active");
  void button.offsetWidth;
  button.classList.add("active");
  copyEWM();
});
document.getElementById("clear-all-btn").addEventListener("click", () => {
  document.getElementById("ticket-num").value = "";
  document.getElementById("multi-ticket").value = "";
});
document.getElementById("scan-ticket").addEventListener("click", () => {
  startBarcodeScan("ticket-num");
});
document.getElementById("cancelScan").addEventListener("click", () => {
  Quagga.stop();
  Quagga.offDetected();
  document.getElementById("scannerModal").style.display = "none";
});

import { initializeApp } from "https://www.gstatic.com/firebasejs/12.18.0/firebase-app.js";
import { getAuth, signInWithEmailAndPassword } from "https://www.gstatic.com/firebasejs/12.18.0/firebase-auth.js";
import {
  getDatabase,
  ref,
  onValue,
  push,
  set,
  get,
  remove,
} from "https://www.gstatic.com/firebasejs/12.18.0/firebase-database.js";

const log = console.log;
const excelInput = document.getElementById("excelInput");
const completeInput = document.getElementById("completeInput");
const cycle = document.getElementById("cycleTable");
// const start = document.getElementById("start-auto");
// const copyBins = document.getElementById("copy-bins");
// const order = document.getElementById("changeOrder");
const tasks = new Map();
const updateLocations = document.getElementById("updateLoc");
const locSubmit = document.getElementById("locSubmit");
const updateMsg = document.getElementById("lastUpdated");
let currentSection;

const firebaseConfig = {
  apiKey: "AIzaSyApvXyYsfG3cW4-Occ1xBz7h1W_lvypaic",
  authDomain: "cycle-count-e9eb7.firebaseapp.com",
  databaseURL: "https://cycle-count-e9eb7-default-rtdb.firebaseio.com",
  projectId: "cycle-count-e9eb7",
  storageBucket: "cycle-count-e9eb7.appspot.com",
  messagingSenderId: "500587496692",
  appId: "1:500587496692:web:5a71aa22f33de95dbc0aa1",
};

const app = initializeApp(firebaseConfig);
const auth = getAuth(app);
const ADMIN_EMAIL = "qDev@gmail.com"
const db = getDatabase(app);
const binRef = ref(db, "bins");

async function loginAdmin(password) {
  return signInWithEmailAndPassword(auth, ADMIN_EMAIL, password);
}

onValue(binRef, (snapshot) => {
  const data = snapshot.val() || {};
  const values = Object.keys(data);

  document.getElementById("section-content").innerHTML = "";
  cycle.innerHTML = "";
  tasks.clear();

  if (values.length > 0) {
    loadSession(values);
  }
  currentSection = null;
  const stamp = new Date().toLocaleString();
  updateMsg.innerText = `Last Updated: ${stamp}`;
});
function generateList(bin) {
  const cycle = document.getElementById("cycleTable");
  const tr = document.createElement("tr");
  tr.setAttribute("id", bin);

  const binContent = document.createElement("td");
  binContent.className = "bin";
  binContent.innerText = bin;

  // const woContent = document.createElement("td");
  // woContent.className = "wo";
  // woContent.innerText = wo;

  // const piContent = document.createElement("td");
  // piContent.innerText = pi;

  tr.appendChild(binContent);
  // tr.appendChild(woContent);
  // tr.appendChild(piContent);

  cycle.appendChild(tr);
}
function displayList(e) {
  if (e.target.classList.contains("sectionBtns")) {
    const section = e.target.innerText;
    currentSection = section;
    const taskSections = tasks.keys();

    cycle.innerHTML = "";

    const thBin = document.createElement("th");
    thBin.innerText = "Bin";

    cycle.appendChild(thBin);

    for (let key of taskSections) {
      const bin = key.slice(0, 3);
      // const val = tasks.get(key);
      if (bin === section) {
        generateList(key);
      }
    }
  }
}
function loadSession(dbValues) {
  excelInput.setAttribute("disabled", "");

  let sectionList = new Set();
  tasks.clear();

  for (let bin of dbValues) {
    const section = bin.slice(0, 3);
    sectionList.add(section);
    tasks.set(bin, true);
  }

  const content = document.getElementById("section-content");
  content.innerHTML = "";

  for (let section of sectionList) {
    const button = document.createElement("button");
    button.innerText = section;
    button.className = "button sectionBtns";
    button.setAttribute("id", section);
    content.appendChild(button);
  }
}
export async function updateDB(value) {
  const bin = value.trim();
  const binRef = ref(db, "bins/" + bin);
  const snapshot = await get(binRef);

  if (snapshot.exists()) return;
  await set(binRef, true);
}
export async function deleteVal(value) {
  const cleanVal = value.trim();
  if (!tasks.has(cleanVal)) return;

  const valRef = ref(db, "bins/" + cleanVal);
  await remove(valRef);
  log(`Removed: ${valRef}`);
}
async function waitForLocPass() {
  const input = document.getElementById("locPass");
  const submit = document.getElementById("locSubmit");

  return new Promise(resolve => {
    submit.addEventListener("click", function handler() {
      submit.removeEventListener("click", handler);
      resolve(input.value);
      input.value = "";
    });
  });
}
async function waitForUpdatePass() {
  const input = document.getElementById("updatePass");
  const submit = document.getElementById("updateSubmit");

  return new Promise(resolve => {
    submit.addEventListener("click", function handler() {
      submit.removeEventListener("click", handler);
      resolve(input.value);
      input.value = "";
    });
  });
}

excelInput.addEventListener("change", async (e) => {
  excelInput.setAttribute("disabled", "");
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
    const section = bin.slice(0, 3);

    sectionList.add(section);
    tasks.set(bin, true);
    updateDB(bin);
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
completeInput.addEventListener("change", async (e) => {
  const password = "qtoolsDev#2077NCSamurai";
  const trigger = document.getElementById("updateTrigger");
  trigger.style.display = "block";

  const answer = await waitForUpdatePass();

  if (answer === password) {
    await loginAdmin(answer);
    trigger.style.display = "none";
  } else return;

  const file = e.target.files[0];
  if (!file) return;

  if (file.name !== "updateCC.xlsx") {
    alert("Invalid file");
    e.target.value = "";
    return;
  }

  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });
  const sheetName = workbook.SheetNames[0];
  const data = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName]);
  data.sort((a, b) => {
    if (a["Storage Bin"] < b["Storage Bin"]) return -1;
    if (a["Storage Bin"] > b["Storage Bin"]) return 1;
    return 0;
  });

  for (let row = 0; row < data.length; row++) {
    if (data[row]["Storage Bin"] === undefined) break;

    const bin = data[row]["Storage Bin"];
    if (tasks.has(bin)) deleteVal(bin);
  }
  location.reload();
});
document.addEventListener("click", (e) => {
  displayList(e);
});
updateLocations.addEventListener("click", () => {
  console.log("Found Me!");
  const trigger = document.getElementById("locTrigger");
  trigger.style.display = "block";
});
locSubmit.addEventListener("click", () => {
  const password = "qtoolsDev#2077NCSamurai";
  const answer = document.getElementById("locPass");
  const trigger = document.getElementById("locTrigger");

  if (answer.value === password) {
    loginAdmin(password);
    answer.value = "";
    excelInput.removeAttribute("disabled", "");
    trigger.style.display = "none";
  } else {
    trigger.style.display = "none";
    answer.value = "";
  } 
});

// Action Bar Functions
// start.addEventListener("click", () => {
//   const warehouseOrders = [...document.querySelectorAll(".wo")].map(
//     (el) => el.innerHTML,
//   );
//   woAutomation(warehouseOrders);
// });
// copyBins.addEventListener("click", () => {
//   copyBins.classList.remove("active");
//   void copyBins.offsetWidth;
//   copyBins.classList.add("active");
//   copyAllBins();
// });
// order.addEventListener("click", () => {
//   const values = [...document.querySelectorAll(".bin")]
//     .map((el) => el.innerHTML).reverse();

//   cycle.innerHTML = "";

//   const thBin = document.createElement("th");
//   thBin.innerText = "Bin";
//   const thWo = document.createElement("th");
//   thWo.innerText = "Warehouse Order";
//   const thPi = document.createElement("th");
//   thPi.innerText = "PI Document";

//   cycle.appendChild(thBin);
//   cycle.appendChild(thWo);
//   cycle.appendChild(thPi);

//   for (let key of values) {
//     const bin = key.slice(0, 3);
//     const val = tasks.get(key);
//     if (bin === currentSection) {
//       generateList(key, val[0], val[1]);
//     }
//   }
// });

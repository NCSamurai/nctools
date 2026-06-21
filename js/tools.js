const pdfInput = document.getElementById("pdfInput");
const downloads = document.getElementById("downloads");
const downloadBtn = document.getElementById("downloadBtn");
let output;

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

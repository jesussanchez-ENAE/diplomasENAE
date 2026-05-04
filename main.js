const { jsPDF } = window.jspdf;

// UI Elements
const dropZone = document.getElementById('drop-zone');
const excelInput = document.getElementById('excel-input');
const progressSection = document.getElementById('progress-section');
const progressBar = document.getElementById('progress-bar');
const statusLabel = document.getElementById('status-label');
const statusCount = document.getElementById('status-count');
const resultsSection = document.getElementById('results-section');
const downloadZipBtn = document.getElementById('download-zip');
const resetBtn = document.getElementById('reset-app');
const downloadSampleBtn = document.getElementById('download-sample');
const pdfSizeSelect = document.getElementById('pdf-size');
const langSelect = document.getElementById('tpl-lang');

// Global Inputs
const creditsInput = document.getElementById('tpl-credits-input');
const courseYearInput = document.getElementById('tpl-course-year-input');
const expeditionDateInput = document.getElementById('tpl-expedition-date-input');

// Signatures
const sigInput1 = document.getElementById('sig-input-1');
const sigName1Input = document.getElementById('sig-name-1');
const sigTitle1Input = document.getElementById('sig-title-1');
const sigInput2 = document.getElementById('sig-input-2');
const sigName2Input = document.getElementById('sig-name-2');
const sigTitle2Input = document.getElementById('sig-title-2');

// Template elements
const diplomaTemplate = document.getElementById('diploma-template');
const tplStudentName = document.getElementById('tpl-student-name');
const tplEventName = document.getElementById('tpl-event-name');
const tplCredits = document.getElementById('tpl-credits');
const tplCourseYear = document.getElementById('tpl-course-year');
const tplLocationDate = document.getElementById('tpl-location-date');

const tplSigName1 = document.getElementById('tpl-sig-name-1-display');
const tplSigTitle1 = document.getElementById('tpl-sig-title-1-display');
const tplSigImg1Actual = document.getElementById('tpl-sig-img-1-display');

const tplSigName2 = document.getElementById('tpl-sig-name-2-display');
const tplSigTitle2 = document.getElementById('tpl-sig-title-2-display');
const tplSigImg2Actual = document.getElementById('tpl-sig-img-2-display');

// Translation Elements
const tplTextCompletion = document.getElementById('tpl-text-completion');
const tplTextCredits = document.getElementById('tpl-text-credits');
const tplTextIssue = document.getElementById('tpl-text-issue');
const tplTextSigned1 = document.getElementById('tpl-text-signed-1');
const tplTextSigned2 = document.getElementById('tpl-text-signed-2');

const translations = {
    en: {
        completion: "Who has successfully completed the degree",
        credits: "In {year}, with a total of {credits} credits, this degree",
        issue: "is issued to certify the application.",
        signed: "Signed:",
        location: "Murcia,"
    },
    es: {
        completion: "Que ha superado con éxito el programa",
        credits: "En {year}, con un total de {credits} créditos, el presente diploma",
        issue: "se expide para certificar la solicitud.",
        signed: "Firmado:",
        location: "Murcia,"
    }
};

let generatedPdfs = [];

// Handle Signature Uploads with UI updates
function setupSignatureHandler(num) {
    const input = document.getElementById(`sig-input-${num}`);
    const actualImg = document.getElementById(`tpl-sig-img-${num}-display`);
    const previewImg = document.getElementById(`tpl-sig-img-${num}`);
    const icon = document.getElementById(`sig-icon-${num}`);
    const label = document.getElementById(`sig-label-${num}`);

    input.addEventListener('change', function(e) {
        const file = e.target.files[0];
        if (file) {
            const reader = new FileReader();
            reader.onload = function(event) {
                // Update Template
                actualImg.src = event.target.result;
                actualImg.style.display = 'block';
                
                // Update UI Preview
                previewImg.src = event.target.result;
                previewImg.style.display = 'block';
                icon.style.display = 'none';
                label.textContent = "Cambiar Firma";
            };
            reader.readAsDataURL(file);
        }
    });
}

setupSignatureHandler(1);
setupSignatureHandler(2);

// Drag & Drop
dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    dropZone.style.borderColor = 'var(--primary)';
});

dropZone.addEventListener('dragleave', () => {
    dropZone.style.borderColor = 'var(--slate-300)';
});

dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    const files = e.dataTransfer.files;
    if (files.length) handleFile(files[0]);
});

excelInput.addEventListener('change', (e) => {
    if (e.target.files.length) handleFile(e.target.files[0]);
});

async function handleFile(file) {
    if (!file.name.match(/\.(xlsx|xls)$/)) {
        alert('Por favor, sube un archivo Excel válido (.xlsx o .xls)');
        return;
    }

    statusLabel.textContent = "Leyendo datos...";
    progressSection.style.display = 'block';
    resultsSection.style.display = 'none';
    progressBar.style.width = '0%';
    
    const reader = new FileReader();
    reader.onload = async (e) => {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        const json = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1 });
        const rows = json.slice(1).filter(row => row[0] || row[1]); 
        if (rows.length === 0) {
            alert('El archivo Excel parece estar vacío.');
            return;
        }
        await processDiplomas(rows, pdfSizeSelect.value);
    };
    reader.readAsArrayBuffer(file);
}

async function processDiplomas(rows, size) {
    generatedPdfs = [];
    const total = rows.length;
    const lang = langSelect.value;
    const trans = translations[lang];

    // Set template size
    if(size === 'a3') {
        diplomaTemplate.style.width = '1587px';
        diplomaTemplate.style.height = '1123px';
    } else {
        diplomaTemplate.style.width = '1123px';
        diplomaTemplate.style.height = '794px';
    }
    
    const dims = size === 'a3' ? [1587, 1123] : [1123, 794];

    // Apply translations
    if (tplTextCompletion) tplTextCompletion.textContent = trans.completion;
    if (tplTextIssue) tplTextIssue.textContent = trans.issue;
    if (tplTextSigned1) tplTextSigned1.textContent = trans.signed;
    if (tplTextSigned2) tplTextSigned2.textContent = trans.signed;

    // Sync static fields
    if (tplSigName1) tplSigName1.textContent = sigName1Input.value;
    if (tplSigTitle1) tplSigTitle1.textContent = sigTitle1Input.value;
    if (tplSigName2) tplSigName2.textContent = sigName2Input.value;
    if (tplSigTitle2) tplSigTitle2.textContent = sigTitle2Input.value;
    
    if (tplLocationDate) tplLocationDate.textContent = `${trans.location} ${expeditionDateInput.value}`;

    for (let i = 0; i < total; i++) {
        const row = rows[i];
        const fullName = `${row[0] || ""} ${row[1] || ""}`.trim().toUpperCase();
        const masterName = row[2] || "Master Program";

        statusLabel.textContent = `Generando: ${fullName}`;
        statusCount.textContent = `${i + 1}/${total}`;
        
        if (tplStudentName) tplStudentName.textContent = fullName;
        if (tplEventName) tplEventName.textContent = masterName;
        
        // Dynamic translation for credits line
        if (tplTextCredits) {
            tplTextCredits.innerHTML = trans.credits
                .replace('{year}', `<span id="tpl-course-year" style="font-weight: bold;">${courseYearInput.value}</span>`)
                .replace('{credits}', `<span id="tpl-credits" style="font-weight: bold;">${creditsInput.value}</span>`);
        }

        const canvas = await html2canvas(diplomaTemplate, { 
            scale: 2, 
            useCORS: true,
            allowTaint: true,
            logging: false
        });
        
        const imgData = canvas.toDataURL('image/jpeg', 0.95);
        const pdf = new jsPDF({ orientation: 'landscape', unit: 'px', format: dims });
        pdf.addImage(imgData, 'JPEG', 0, 0, dims[0], dims[1]);
        generatedPdfs.push({ name: `${fullName.replace(/\s+/g, '_')}.pdf`, blob: pdf.output('blob') });
        progressBar.style.width = `${((i + 1) / total) * 100}%`;
    }
    statusLabel.textContent = "¡Completado!";
    resultsSection.style.display = 'block';
}

downloadZipBtn.addEventListener('click', async () => {
    const zip = new JSZip();
    generatedPdfs.forEach(p => zip.file(p.name, p.blob));
    const content = await zip.generateAsync({ type: "blob" });
    saveAs(content, "Diplomas_Oficiales_ENAE.zip");
});

resetBtn.addEventListener('click', () => {
    progressSection.style.display = 'none';
    resultsSection.style.display = 'none';
    excelInput.value = '';
});

downloadSampleBtn.addEventListener('click', (e) => {
    e.preventDefault();
    const data = [["Nombre", "Apellidos", "Máster"], ["John", "Doe", "Master in Management"], ["Jane", "Smith", "Executive MBA"]];
    const ws = XLSX.utils.aoa_to_sheet(data);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, "Alumnos");
    XLSX.writeFile(wb, "Ejemplo_ENAE.xlsx");
});

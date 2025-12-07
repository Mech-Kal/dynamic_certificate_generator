// Modal functions
const modal = document.getElementById("howToUseModal");

function openModal() { modal.style.display = "block"; }
function closeModal() { modal.style.display = "none"; }
window.onclick = function(event) {
    if (event.target == modal) modal.style.display = "none";
}

// Parameters logic
let parameters = [];

function addParameter() {
    const input = document.getElementById("paramInput");
    let val = input.value.trim();
    if(val && !parameters.includes(val.toLowerCase())){
        parameters.push(val.toLowerCase());
        const li = document.createElement("li");
        li.textContent = val.toLowerCase();
        document.getElementById("paramList").appendChild(li);
    }
    input.value = "";
}

function completeParameters() {
    if(parameters.length === 0){
        alert("Add at least one parameter");
        return;
    }
    document.getElementById("step1").style.display = "none";
    document.getElementById("step2").style.display = "block";
    document.getElementById("paramDisplay").textContent = parameters.join(", ");
}

// Logging function
function logMessage(msg){
    const log = document.getElementById("log");
    log.textContent += msg + "\n";
    log.scrollTop = log.scrollHeight;
}

// Generate Certificates
async function generateCertificates() {
    const templateFile = document.getElementById("template").files[0];
    const excelFile = document.getElementById("excel").files[0];

    if(!templateFile || !excelFile){
        alert("Upload both DOCX template and Excel file");
        return;
    }

    logMessage("Starting...");

    const templateBuffer = await templateFile.arrayBuffer();
    const excelBuffer = await excelFile.arrayBuffer();
    const workbook = XLSX.read(excelBuffer,{type:"array"});
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(sheet);

    logMessage("Rows found: " + rows.length);

    rows.forEach((row,i)=>{
        const data = {};
        parameters.forEach(p=>{
            data[p] = row[p] || "";
        });

        const zip = new PizZip(templateBuffer);
        const doc = new docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
        doc.setData(data);

        try { doc.render(); }
        catch(error){
            logMessage(`Row ${i+1} render error: ${error}`);
            return;
        }

        const output = doc.getZip().generate({type:"blob"});
        const link = document.createElement("a");
        link.href = URL.createObjectURL(output);
        link.download = `${data.name || "certificate"}_certificate.docx`;
        link.click();

        logMessage(`Certificate generated for row ${i+1}`);
    });

    logMessage("All certificates generated!");
}

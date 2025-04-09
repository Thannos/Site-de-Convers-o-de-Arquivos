let selectedFiles = [];

function selectFiles() {
    document.getElementById("file-input").click();
}

function handleFiles(files) {
    const fileList = document.getElementById("file-list");
    const buttons = document.getElementById("buttons-container");

    for (let file of files) {
        if (!selectedFiles.some(f => f.name === file.name)) {
            selectedFiles.push(file);
            let fileDiv = document.createElement("div");
            fileDiv.classList.add("file-card");
            fileDiv.dataset.filename = file.name;
            fileDiv.innerHTML = `
                <button class="remove-btn" onclick="removeFile('${file.name}')">X</button>
                <p>${file.name}</p>
            `;
            fileList.appendChild(fileDiv);
        }
    }
    buttons.style.display = selectedFiles.length > 0 ? "block" : "none";
}

function dropHandler(event) {
    event.preventDefault();
    handleFiles(event.dataTransfer.files);
}

function dragoverHandler(event) {
    event.preventDefault();
}

function removeFile(fileName) {
    selectedFiles = selectedFiles.filter(file => file.name !== fileName);
    document.querySelector(`.file-card[data-filename='${fileName}']`).remove();
    document.getElementById("buttons-container").style.display = selectedFiles.length > 0 ? "block" : "none";
}

function sortFiles() {
    selectedFiles.sort((a, b) => a.name.localeCompare(b.name));
    document.getElementById("file-list").innerHTML = "";
    handleFiles(selectedFiles);
}

async function convertFiles() {
    for (let file of selectedFiles) {
        let reader = new FileReader();
        reader.onload = async function(event) {
            let result = await mammoth.extractRawText({ arrayBuffer: event.target.result });
            let rows = result.value.split('\n').map(line => line.split(';'));
            let ws = XLSX.utils.aoa_to_sheet(rows);
            let wb = XLSX.utils.book_new();
            XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
            XLSX.writeFile(wb, file.name.replace(/\.[^.]+$/, "") + ".xlsx");
        };
        reader.readAsArrayBuffer(file);
    }
}

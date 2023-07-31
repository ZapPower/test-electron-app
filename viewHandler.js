const typeChoose = document.getElementById('typeSelect');
const label = document.getElementById('excelLabel');
const uploadButton = document.getElementById('excelButton');

// Change messaging with different selection
typeChoose.addEventListener('change', () => {
    var sheetType = typeChoose.value;
    if (sheetType == "court") {
        label.innerText = "Import excel sheet here:";
        uploadButton.innerText = "Choose Excel File";
    } else {
        label.innerText = "Import CSV sheet here:";
        uploadButton.innerText = "Choose CSV File";
    }
});
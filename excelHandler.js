const XLSX = require('xlsx');
const {ipcRenderer} = require('electron');
const { jsPDF } = require('jspdf');

// loads and parses excel file using xlsx
function loadexcel() {
    // request the file from the main process
    ipcRenderer.send('requestExcelFile');

    // when file is sent back, open and parse it into JSON
    ipcRenderer.on('excelFile', (event, filepath) => {
        // create workbook
        var workbook = XLSX.read(filepath, {type: 'file', cellDates: true});
        var sheet = workbook.SheetNames[0];
        // display sheet name
        document.getElementById('sheetname').innerText = sheet;
        // convert to JSON. Parse as string for easier formatting
        var sheetJSON = XLSX.utils.sheet_to_json(workbook.Sheets[sheet], {raw: false});

        // get new dictionary version
        calendarDict = getAppearanceDict(sheetJSON);
        // apply HTML version of new dictionary
        document.getElementById('calendar').innerHTML = createAppearanceHTML(calendarDict);
        // filter and style the applied HTML
        filterAndStyle();
        // TODO: after finishing styling, generate PDF version of table to be downloaded by client

        // un-hide download PDF button
        document.getElementById('createPDF').style.visibility = 'visible';
    });   
}

// creates and returns a new dictionary with only the necessary data obtained from excel sheet
function getAppearanceDict(sheet) {
    calendar = [];
    for (var row = 0; row < sheet.length; row++) {
        calendar.push({});
        calendar[row]['File No.'] = sheet[row]['Client Matter'];
        calendar[row]['Case Name'] = sheet[row]['Case Name'];
        calendar[row]['Adj\'ed?'] = " ";
        calendar[row]['Room/Part'] = sheet[row]['Part/Room'];
        calendar[row]['Time'] = sheet[row]['Appearance Time'];
        calendar[row]['Detail'] = sheet[row]['Type'];
        calendar[row]['County'] = sheet[row]['County'];
        calendar[row]['Date'] = sheet[row]['Appearance Date'];
    }
    return calendar;
}

// creates and returns an HTML table consisting of the given dictionary using SheetJS
function createAppearanceHTML(calendar) {
    newSheet = XLSX.utils.json_to_sheet(calendar);
    table = XLSX.utils.sheet_to_html(newSheet);
    return table;
}

/*
TODO: set size for each element as needed (add CSS classes to each column as needed to specify) 
ex. 
HTML:
class="date"
CSS:
.date {
    width: 10%;
}
*/
// filters the auto-generated HTML from SheetJS & styles the generated table
function filterAndStyle() {
    // get table elements
    var elements = document.getElementsByTagName('td');
    elements[7].remove();
    elements[6].remove();
    var dateStor = undefined;
    var courtStor = undefined;
    for (var i = 0; i < elements.length; i++) {
        elem = elements[i];
        // check if element is in first row
        if (elem.id.slice(-1) == "1" && elem.id.length == 6) {
            // make it a header
            elem.outerHTML = `<th id="${elem.id}">${elem.innerText}</th>`;
            // decrement counter as this operation removes element from the list
            i--;
        }

        // check element is date
        if (elem.id.charAt(4) == "H") {
            // check if stored date matches. If not push new element at stored index & delete
            if (elem.innerText != dateStor) {
                dateStor = elem.innerText;
                newParent = document.createElement("tr");
                newParent.innerHTML = `<th id="${elem.id}" colspan=6>${elem.innerText}</th>`;
                elements[i].parentNode.parentNode.insertBefore(newParent, elements[i].parentNode);
            }
            elements[i].remove();
            i--;
        }

        // check if elem is court
        if (elem.id.charAt(4) == "G") {
            // check if stored court matches. If not push new element at previous index & delete
            if (elem.innerText != courtStor) {
                courtStor = elem.innerText;
                newParent = document.createElement("tr");
                newParent.innerHTML = `<th id="${elem.id}" colspan=6>${elem.innerText}</th>`;
                elements[i].parentNode.parentNode.insertBefore(newParent, elements[i].parentNode);
            }
            elements[i].remove();
            i--;
        }
    }
}

// generate downloadable PDF & prompt client to save
function generatePDF() {
    ipcRenderer.send('requestDownloadPath');

    ipcRenderer.on('downloadPath', (event, filepath) => {
        var doc = new jsPDF({
            orientation: 'l'
        });
        var elem = document.querySelector("#calendar");
        doc.html(elem, {
            callback: function (doc) {
                doc.save(filepath);
                alert("File saved to " + filepath);
            },
            html2canvas: {
                scale: 0.4
            }
        });
    });
}
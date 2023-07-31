const XLSX = require('xlsx');
const { ipcRenderer } = require('electron');
const { jsPDF } = require('jspdf');

// loads and parses excel file using xlsx
function loadexcel() {
    // get filetype (EBT or Court) to know how to parse & filter
    var filetype = document.getElementById('typeSelect').value;

    // request the file from the main process
    ipcRenderer.send('requestExcelFile', filetype);

    // when file is sent back, open and parse it into JSON
    ipcRenderer.on('excelFile', (event, filepath) => {
        // create workbook
        var workbook = XLSX.read(filepath, {type: 'file', cellDates: true});
        var sheet = workbook.SheetNames[0];
        // display sheet name
        document.getElementById('sheetname').innerText = sheet;
        // convert to JSON. Parse as string for easier formatting
        var sheetJSON = XLSX.utils.sheet_to_json(workbook.Sheets[sheet], {raw: false});
        console.log(sheetJSON);

        // get new dictionary version
        if (filetype == 'court') {
            calendarDict = getAppearanceDictCourt(sheetJSON);
        } else {
            calendarDict = getAppearanceDictEBT(sheetJSON);
        };
        // apply HTML version of new dictionary
        document.getElementById('calendar').innerHTML = createAppearanceHTML(calendarDict);
        // filter and style the applied HTML
        filterAndStyle(filetype);
        // un-hide download PDF button
        document.getElementById('createPDF').style.visibility = 'visible';
    });   
}

// creates and returns a new dictionary with only the necessary data obtained from excel sheet (for type Court)
function getAppearanceDictCourt(sheet) {
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
        calendar[row][''] = "";
    };
    return calendar;
}

// creates and returns a new dictionary with only the necessary data obtained from excel sheet (for type EBT)
function getAppearanceDictEBT(sheet) {
    calendar = [];
    for (var row = 0; row < sheet.length; row++) {
        calendar.push({});
        calendar[row]['Subject'] = sheet[row]['Subject'];
        calendar[row]['Start Date'] = sheet[row]['Start Date'];
        calendar[row]['Start Time'] = sheet[row]['Start Time'];
        calendar[row]['Required Attendees'] = sheet[row]['Required Attendees'];
        calendar[row]['Location'] = sheet[row]['Location'];
        calendar[row]['Notes'] = "";
        calendar[row][''] = "";
    };
    return calendar;
}

// <button aria-label='delete item' type='button' contenteditable='false' class='delButton' id='del${delCount}'>X</button>

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
function filterAndStyle(filetype) {
    // get table elements
    var elements = document.getElementsByTagName('td');
    if (filetype == "court") {
        elements[7].remove();
        elements[6].remove();
    } else {
        elements[1].remove();
    };
    var dateStor = undefined;
    var courtStor = undefined;

    // choose column to check date based on filetype
    var delcol = "G";
    var col = "H";
    var colspan = 6;
    if (filetype == "ebt") {
        col = "B";
        colspan = 5;
    };

    for (var i = 0; i < elements.length; i++) {
        elem = elements[i];
        // check if element is in first row
        if (elem.id.slice(-1) == "1" && elem.id.length == 6) {
            // make it a header
            if (elem.id.charAt(4) != delcol) {
                elem.outerHTML = `<th id="${elem.id}">${elem.innerText}</th>`;
            } else {
                elem.outerHTML = `<th id="${elem.id}" class="delCol">${elem.innerText}</th>`;
            }
            // decrement counter as this operation removes element from the list
            i--;
        };

        // check element is date
        if (elem.id.charAt(4) == col) {
            // check if stored date matches. If not push new element at stored index & delete
            if (elem.innerText != dateStor) {
                dateStor = elem.innerText;
                newParent = document.createElement("tr");
                newParent.innerHTML = `<th id="${elem.id}" colspan=${colspan}>${elem.innerText}</th>`;
                elements[i].parentNode.parentNode.insertBefore(newParent, elements[i].parentNode);
            }
            elements[i].remove();
            i--;
        };

        // check if elem is court
        if (filetype == "court" && elem.id.charAt(4) == "G") {
            // check if stored court matches. If not push new element at previous index & delete
            if (elem.innerText != courtStor) {
                courtStor = elem.innerText;
                newParent = document.createElement("tr");
                newParent.innerHTML = `<th id="${elem.id}" colspan=${colspan}>${elem.innerText}</th>`;
                elements[i].parentNode.parentNode.insertBefore(newParent, elements[i].parentNode);
            }
            elements[i].remove();
            i--;
        };

        //add delCol class to elements in delete column
        if (elem.id.charAt(4) == delcol) {
            elem.className = "delCol";
        }
    }

    // create calendar header
    const header = document.createElement("tr");
    headerText = "Court";
    if (filetype == "ebt") {
        headerText = "EBT";
    };
    header.innerHTML = `<th id="calHead" colspan=${colspan + 1} class="calendarHead">${headerText}  Calendar</th>`;
    elem = document.getElementsByTagName("tr")[0];
    elem.parentNode.insertBefore(header, elem);

    // fill in delete column and add in delete buttons
    var newChild;
    elements = document.getElementsByTagName('tr');
    for (var row = 2; row < elements.length; row++) {
        elem = elements[row];
        elem.id = "row"+row;
        if (elem.childElementCount == 1) {
            newChild = document.createElement("td");
            newChild.className = "delCol";
            newChild.innerHTML = `<button aria-label='delete item' type='button' contenteditable='false' class='delButton' id='del${row}' onclick='delRow(${row})'>X</button>`;
            elem.appendChild(newChild);
        } else {
            elem.lastChild.innerHTML = `<button aria-label='delete item' type='button' contenteditable='false' class='delButton' id='del${row}' onclick='delRow(${row})'>X</button>`;
        };
    };
}

// generate downloadable PDF & prompt client to save
function generatePDF() {
    // get filetype
    var filetype = document.getElementById('typeSelect').value;
    var pdfName = "Molod Spitz & DeSantis, PC Court Appearance Report.pdf";
    if (filetype == "ebt") {
        pdfName = "Molod Spitz & DeSantis, PC EBT Appearance Report.pdf";
    };

    ipcRenderer.send('requestDownloadPath', pdfName);

    ipcRenderer.on('downloadPath', (event, filepath) => {
        var doc = new jsPDF({
            orientation: 'l'
        });
        var elem = document.querySelector("#calendar");
        // download pdf to chosen location
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

// delete row function for row buttons
function delRow(row) {
    document.getElementById("row"+row).remove();
}
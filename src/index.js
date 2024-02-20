
// Do bindings here
let printSection = document.getElementById('printSection');
let processButton = document.getElementById('processButton');
let excelFilePicker = document.getElementById('excelFilePicker');
let monthPicker = document.getElementById('monthPicker');
let infoText1 = document.getElementById('infoText1');
let infoText2 = document.getElementById('infoText2');
let infoText3 = document.getElementById('infoText3');
let infoMonth = document.getElementById('infoMonth');


// other global variables
let excelFileData;
let actualMonth = monthPicker.value;


infoText1.innerText = "Noch keine Datei ausgewählt";


// excel file picker change event
excelFilePicker.addEventListener('change', function(e) {
    readFile(excelFilePicker.files[0]);
    infoText1.innerText = "Ausgewählte Datei - Stand " + new Date().toLocaleTimeString()+ " : " + excelFilePicker.files[0].name; 
});


// month picker change event
monthPicker.addEventListener('change', function(e) {
    actualMonth = e.target.value;
});


// button click event
processButton.addEventListener('click', function() {
    
    let workbook;
    try{
        workbook = XLSX.read(excelFileData, {type: 'array'});
    }
    catch (error){
        infoText1.innerText = "BITTE EXCEL DATEI AUSWÄHLEN";
        return;
    }

    // Use the workbook here
    const firstSheetName = workbook.SheetNames[0];
    infoText2.innerText = "Erstes Tabellenblatt: " + firstSheetName;

    const worksheet = workbook.Sheets[firstSheetName];
    const actualDescription = worksheet['A1'].v;
    infoText3.innerText = "Inhalt Zelle A1: " + actualDescription;
    
    infoMonth.innerText = "Extrahierter Monat: " + actualMonth;

    const jsonData = XLSX.utils.sheet_to_json(worksheet);

    //console.log(jsonData);

    // to collect the html as a string before adding it to the innerHTML of the printSection
    let tableHTML = "";
    
    tableHTML += "<table>";

    jsonData.forEach(item => {

        //console.log(item);
        if (item.hasOwnProperty(actualMonth)) {
            
            tableHTML += "<tr>";
            let value = ""; // by default print nothing for value

            if (item.hasOwnProperty("meta")){
                                                
                switch (item["meta"]) {
                    case "header":
                        tableHTML += "</tr></table><table><tr>";
                        tableHTML += '<td class="header">' + item[actualDescription] + "</td>";
                        if (item[actualMonth] != actualMonth){
                            // if the value is not the name of the month, take it as an amount of money
                            value = parseFloat(item[actualMonth]).toFixed(2) + " Euro" 
                        }                        
                        tableHTML += '<td class="headerValue">' + value + "</td>";
                        tableHTML += '<td class="mark"></td>';
                        break;
                    case "item":
                        tableHTML += '<td class="item">' + item[actualDescription] + "</td>";
                        tableHTML += '<td class="itemValue">' + parseFloat(item[actualMonth]).toFixed(2) + " Euro" + "</td>";
                        tableHTML += '<td class="mark"></td>';
                        break;
                    case "subheader":
                        tableHTML += '<td class="subHeader">' + item[actualDescription] + "</td>"; 
                        if (item[actualMonth] != actualMonth){
                            // if the value is not the name of the month, take it as an amount of money
                            value = parseFloat(item[actualMonth]).toFixed(2) + " Euro" 
                        }  
                        tableHTML += '<td class="subHeaderValue">' + value + "</td>"; 
                        tableHTML += '<td class="mark"></td>';
                        break;
                    default:
                        tableHTML += '<td class="item">wrong meta</td>';
                        tableHTML += '<td class="itemValue">wrong meta</td>';
                        tableHTML += '<td class="mark"></td>';
                        break;
                }  
                
            }
            else{
                tableHTML += '<td class="item">no meta field</td>';
                tableHTML += '<td class="itemValue">no meta field</td>';
                tableHTML += '<td class="mark"></td>';
            }
            
            tableHTML += "</tr>";

        }

    });

    tableHTML += "</table>";

    printSection.innerHTML = tableHTML;
    //console.log(tableHTML);

});


// Function to read the file
function readFile(file) {
    const reader = new FileReader();
    reader.onload = function(e) {
        excelFileData = new Uint8Array(e.target.result);
    };
    reader.readAsArrayBuffer(file);
}
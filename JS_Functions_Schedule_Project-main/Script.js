
//Naryan Sambhi - function to check if content is valid / used cells 

// Function to check if the content is a month
function isMonth(content) {
    const months = [
        'January', 'February', 'March', 'April', 'May', 'June', 'July',
        'August', 'September', 'October', 'November', 'December'
    ];

    return months.includes(content);
}

//Function to check if content is a shift
function isShift(content) {
    const Shifts = [
        '', 'DO', 'NO', 'NT', 'DT', 'Dt', 'Nt',
        'NoT', 'DoT', 'D7', 'D7T', 'D7t', 'O', 'Mt', 'OC', 'H', 'RH', 'NOt', 'DOt'
    ];

    return Shifts.includes(content);
}


// Sort By Specific Column
document.getElementById("myForm").addEventListener("submit", function (event) {
    event.preventDefault(); // Prevent the form from submitting normally

    // Get user input from the form
    var day = document.getElementById("day").value;

    // Validation
    if (day < 1 || day > 38) {
        alert("Invalid input. Please select a valid day");
        return;
    }

    // Get the table element
    const table = document.getElementsByTagName('table')[0];

    // Loop through each row in the table
    for (let i = 0; i < table.rows.length; i++) {

        //get content 
        const cellToCheck = table.rows[i].cells[day];

        // Short form trim
        const Content = cellToCheck.innerHTML.trim();

        // Check if the cell is empty or contains specific values

        //dummy row  - fake cells for previous months date (should hide as it will be empty)
        if (cellToCheck.style.backgroundColor === 'rgb(166, 166, 166)' || cellToCheck.style.backgroundColor === '#A6A6A6') {
            table.rows[i].style.display = "none"; // Hide the entire row
        }
        //content row
        if (!isNaN(Content) || isMonth(Content) || isShift(Content)) {
            cellToCheck.style.backgroundColor = "#9ffcfc";
            //hide row
        } else {
            table.rows[i].style.display = "none"; // Hide the entire row
        }
    }
});



//Naryan Sambhi - 2023 - expand table and remove invisble values

function hideLastSevenCellsAndExpandTable() {
    const tableElements = document.getElementsByTagName('table');

    for (let i = 0; i < tableElements.length; i++) {
        const rows = tableElements[i].querySelectorAll('tr'); // Select all rows in the current table

        rows.forEach((row) => {
            const cells = Array.from(row.querySelectorAll('td')); // Convert NodeList to an array
            const lastSevenCells = cells.slice(-7); // Get the last seven cells of the row

            lastSevenCells.forEach((cell) => {
                // Check if the cell's content is "DT" or other headers before hiding it

                const cellContent = cell.textContent.trim();
                if (cellContent !== "DT" && cellContent !== "N" && cellContent !== "NT" && cellContent !== "OP%") {
                    cell.style.display = 'none';
                }



            });
        });

        // Set the width of the table to 1250pt
        tableElements[i].style.width = '1250pt';
    }
}

document.addEventListener('DOMContentLoaded', hideLastSevenCellsAndExpandTable);



//Naryan sambhi - create new columns for desired usage 


//read rows function with inputed values, returns sum if found
function readRows(row, ...values) {


    //defines
    const table = document.getElementsByTagName('tr')[row];

    const columns = 38;
    let sum = 0;

    // go through cells 
    for (let i = 1; i < columns; i++) {
        const cell = table.cells[i];

        // Skip invalid or unrelated cells
        if (!cell || cell.tagName === 'TH' || table.style.display === 'none') {
            continue;
        }

        // Check content within the cell for desired contents

        const text = cell.textContent;

        if (values.includes(text)) {
            sum++;
        }
    }

    return sum;
}

//adds actual columns to table 

function addColumnToTable(tableId, columnIndex, headerText, ...values) {

    //defines 
    const table = document.getElementsByTagName('table')[0];

    // Apply header 
    const headerCell = table.rows[0].cells[columnIndex];

    headerCell.innerHTML = headerText;
    headerCell.style.backgroundColor = '#d8e4bc';
    headerCell.style.fontSize = '15px';
    headerCell.style.border = '1.5px solid black';
    headerCell.style.textAlign = 'center';
    headerCell.style.verticalAlign = 'middle';
    headerCell.style.color = 'black';



    //create columns
    for (let i = 1; i < table.rows.length; i++) {
        const row = table.rows[i];

        // Skip the row if the cell is invalid
        if (!row.cells[columnIndex]) {
            continue;
        }

        //insert cell and cell values 
        const cell = row.insertCell(columnIndex);

        cell.style.fontSize = '15px';
        cell.style.border = '1.5px solid black';
        cell.style.backgroundColor = '#E4DFEC';
        cell.style.textAlign = 'center';
        cell.style.verticalAlign = 'middle';
        cell.style.color = 'black';


        const sum = readRows(i, ...values);
        cell.innerHTML = sum;
    }
}



function addColumnToTablePercent(tableId, columnIndex) {
    // Get the table element by its ID
    const table = document.getElementsByTagName('table')[0];

    // Apply header
    const headerCell = table.rows[0].insertCell(columnIndex);

    headerCell.innerHTML = 'OP%';
    headerCell.style.backgroundColor = '#d8e4bc';
    headerCell.style.fontSize = '15px';
    headerCell.style.border = '1.5px solid black';
    headerCell.style.textAlign = 'center';
    headerCell.style.verticalAlign = 'middle';


    // Calculate and insert values for each row
    for (let i = 1; i < table.rows.length; i++) {
        const row = table.rows[i];

        // Skip the row if the cell is invalid
        if (!row.cells[columnIndex]) {
            continue;
        }

        // Read cell values for calculations
        var sum1 = readRows(i, 'D', 'D', 'Do', 'DO', 'D7',);
        var sum2 = readRows(i, 'N', 'N', 'No', 'NO', 'N7',);

        var sum3 = readRows(i, 'DT', 'DT', 'Dt', 'DOt', 'DOT', 'D7t', 'D7T',);
        var sum4 = readRows(i, 'NT', 'NT', 'Nt', 'NOt', 'NOT', 'N7t', 'N7T',);

        var InPerson = sum1 + sum2;
        var total = sum1 + sum2 + sum3 + sum4;

        // Calculate percentage as an integer
        var percentage = Math.round((InPerson / total) * 100);

        // Insert the percentage value into the cell
        const cell = row.insertCell(columnIndex);

        cell.innerHTML = percentage + '%';


        cell.style.fontSize = '15px';
        cell.style.border = '1.5px solid black';
        cell.style.backgroundColor = '#E4DFEC';
        cell.style.textAlign = 'center';
        cell.style.verticalAlign = 'middle';

        if (percentage < 60) {
            cell.style.color = 'red';
        }

    }
}



addColumnToTable('myTable', 38, 'D', 'D', 'Do', 'DO', 'D7', '', '', '');
addColumnToTable('myTable', 39, 'DT', 'DT', 'Dt', 'DOt', 'DOT', 'D7t', 'D7T', '');
addColumnToTable('myTable', 40, 'N', 'N', 'No', 'NO', 'N7', '', '', '');
addColumnToTable('myTable', 41, 'NT', 'NT', 'Nt', 'NOt', 'NOT', 'N7t', 'N7T', '');

addColumnToTablePercent('myTable', 42);

var TotalColumns = 5; //important for sorting and filtering - tells program to skip new columns and find correct filters


// checks all dropdown filters every time any is changed, displays the employee that pass all filters 
function filterPage() {

    //Naryan Sambhi - 2023 - checks if existing table contains new totals rows, if so remove them before filtering again.
    function check_table() {

        //get table
        let table = document.getElementsByTagName('tr');


        //define
        var rows = table.length;

        var cell = table[0].cells[0];

        var text = cell.textContent;


        //check first row, first column, first value
        if (text.includes('Total')) {

            document.getElementsByTagName("tr")[0].remove();
            document.getElementsByTagName("tr")[0].remove();


        }

    }

    check_table();


    // Naryan Sambhi - Student - 2023 - sort by columns function when page is filtered 

    //read columns and return sum of all contents with column
    function readColumns(column, ...values) {
        const table = document.getElementsByTagName('tr');
        const rows = table.length;
        let sum = 0;

        for (let i = 1; i < rows; i++) {
            const cell = table[i].cells[column];

            //stop garbage
            if (!cell || cell.tagName === 'TH' || table[i].style.display === 'none') {
                continue;
            }

            const text = cell.textContent.trim();
            if (values.includes(text)) {
                sum++;
            }
        }

        return sum;
    }

    function createAndPopulateTable(shiftType, values, headerText, colorThreshold) {
        const table = document.getElementsByTagName('table')[0];
        const newRow = table.insertRow(0);

        // First cell / label
        const firstValue = `${headerText} Shifts: `;
        const firstCell = newRow.insertCell(0);
        firstCell.innerHTML = firstValue;


        firstCell.innerHTML = headerText;
        firstCell.style.backgroundColor = '#d8e4bc';
        firstCell.style.fontSize = '15px';
        firstCell.style.border = '1.5px solid black';
        firstCell.style.textAlign = 'center';
        firstCell.style.verticalAlign = 'middle';
        firstCell.style.color = 'black';



        // Create table and populate with its columns sum
        for (let i = 1; i < 38; i++) {
            const cell = newRow.insertCell(i);
            cell.className = 'cell';
            const sum = readColumns(i, ...values);
            cell.innerHTML = sum;

            if (sum < colorThreshold) {
                cell.style.color = 'red';
            }

            cell.style.fontSize = '15px';
            cell.style.border = '1.5px solid black';
            cell.style.backgroundColor = '#E4DFEC';
            cell.style.textAlign = 'center';
            cell.style.verticalAlign = 'middle';

            cell.style.display = 'table-cell';
        }
    }

    // Naryan Sambhi 2023 - Create dayshift table and populate
    createAndPopulateTable('Day', ['D', 'Dt', 'DO', 'Dot', 'DT', 'DOT', 'DOt'], 'Total Day', 1);

    // Naryan Sambhi 2023 - Create nightshift table and populate
    createAndPopulateTable('Night', ['N', 'Nt', 'NO', 'Not', 'NT', 'NOT', 'NOt'], 'Total Night', 2);




    window.scroll(0, 0);

}




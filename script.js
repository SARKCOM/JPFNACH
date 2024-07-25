let excelData = [];
const correctPassword = 'jpf123'; // Set your password here

function checkPassword() {
    const passwordInput = document.getElementById('password-input').value;
    if (passwordInput === correctPassword) {
        document.getElementById('password-container').style.display = 'none';
        document.getElementById('main-container').style.display = 'block';
        loadExcelData(); // Load Excel data after the password is confirmed
    } else {
        alert('You are not authorized. Please check the password.');
    }
}

function loadExcelData() {
    fetch('data.xlsx')
        .then(response => {
            if (!response.ok) {
                throw new Error(`HTTP error! status: ${response.status}`);
            }
            return response.arrayBuffer();
        })
        .then(data => {
            const workbook = XLSX.read(data, { type: 'array' });
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            excelData = XLSX.utils.sheet_to_json(firstSheet, { header: 1 });

            console.log('Excel Data:', excelData); // Debugging
            populateColumnSelect();
        })
        .catch(error => {
            console.error('Error fetching or processing the Excel file:', error); // Debugging
        });
}

function populateColumnSelect() {
    const columnSelect = document.getElementById('column-select');
    const headers = excelData[0];

    headers.slice(0, 10).forEach((header, index) => {
        const option = document.createElement('option');
        option.value = index;
        option.textContent = header;
        columnSelect.appendChild(option);
    });

    console.log('Headers:', headers.slice(0, 10)); // Debugging
}

function searchData() {
    const columnSelect = document.getElementById('column-select');
    const searchInput = document.getElementById('search-input');
    const searchTerm = searchInput.value.toLowerCase();
    const columnIndex = parseInt(columnSelect.value, 10); // Ensure column index is an integer

    console.log('Selected Column Index:', columnIndex); // Debugging
    console.log('Search Term:', searchTerm); // Debugging

    const results = excelData.slice(1).filter(row => {
        if (row[columnIndex] !== undefined) {
            return row[columnIndex].toString().toLowerCase().includes(searchTerm);
        }
        return false;
    });

    console.log('Search Results:', results); // Debugging
    displayResults(results, columnIndex);
}

function displayResults(results, columnIndex) {
    const resultsDiv = document.getElementById('results');
    resultsDiv.innerHTML = '';

    if (results.length === 0) {
        resultsDiv.textContent = 'No results found';
        return;
    }

    results.forEach(result => {
        const horizontalTable = document.createElement('table');
        const horizontalThead = document.createElement('thead');
        const horizontalTbody = document.createElement('tbody');
        const headerRow = document.createElement('tr');
        const dataRow = document.createElement('tr');

        // Create header row and data row for the first 10 columns
        result.slice(0, 10).forEach((cell, index) => {
            const th = document.createElement('th');
            const td = document.createElement('td');

            th.textContent = excelData[0][index];

            if (index === columnIndex && typeof cell === 'number' && isExcelDate(cell)) {
                td.textContent = convertExcelDate(cell);
            } else {
                td.textContent = cell;
            }

            headerRow.appendChild(th);
            dataRow.appendChild(td);
        });

        horizontalThead.appendChild(headerRow);
        horizontalTbody.appendChild(dataRow);
        horizontalTable.appendChild(horizontalThead);
        horizontalTable.appendChild(horizontalTbody);

        resultsDiv.appendChild(horizontalTable);

        // Add "Description" line
        const descriptionLine = document.createElement('div');
        descriptionLine.textContent = 'Description';
        descriptionLine.style.marginTop = '10px';
        descriptionLine.style.marginBottom = '10px';
        descriptionLine.style.fontWeight = 'bold';
        resultsDiv.appendChild(descriptionLine);

        // Add "Details" and "Remarks" text
        const paymentInfoDiv = document.createElement('div');
        paymentInfoDiv.style.display = 'flex';
        paymentInfoDiv.style.justifyContent = 'space-between';

        const details = document.createElement('div');
        details.textContent = 'Details';
        details.style.fontWeight = 'bold';

        const remarks = document.createElement('div');
        remarks.textContent = 'Remarks';
        remarks.style.fontWeight = 'bold';

        paymentInfoDiv.appendChild(details);
        paymentInfoDiv.appendChild(remarks);
        resultsDiv.appendChild(paymentInfoDiv);

        // Create vertical table for the remaining columns
        const verticalTable = document.createElement('table');
        const verticalThead = document.createElement('thead');
        const verticalTbody = document.createElement('tbody');

        const verticalHeaders = ['Description', 'Details', 'Remarks'];

        verticalHeaders.forEach(header => {
            const verticalTh = document.createElement('th');
            verticalTh.textContent = header;
            verticalThead.appendChild(verticalTh);
        });

        result.slice(10).forEach((cell, index) => {
            const verticalTr = document.createElement('tr');
            const descriptionCell = document.createElement('td');
            const detailsCell = document.createElement('td');
            const remarksCell = document.createElement('td');

            descriptionCell.textContent = verticalHeaders[index % 3];
            detailsCell.textContent = cell;
            remarksCell.textContent = cell;

            verticalTr.appendChild(descriptionCell);
            verticalTr.appendChild(detailsCell);
            verticalTr.appendChild(remarksCell);

            verticalTbody.appendChild(verticalTr);
        });

        verticalTable.appendChild(verticalThead);
        verticalTable.appendChild(verticalTbody);

        resultsDiv.appendChild(verticalTable);
    });
}

function isExcelDate(num) {
    return num > 25569; // Simple check to see if the number is an Excel date
}

function convertExcelDate(num) {
    const date = new Date((num - 25569) * 86400 * 1000);
    const yyyy = date.getFullYear();
    const mm = ('0' + (date.getMonth() + 1)).slice(-2);
    const dd = ('0' + date.getDate()).slice(-2);
    return `${yyyy}-${mm}-${dd}`;
}

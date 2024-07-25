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

    headers.forEach((header, index) => {
        const option = document.createElement('option');
        option.value = index;
        option.textContent = header;
        columnSelect.appendChild(option);
    });

    console.log('Headers:', headers); // Debugging
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
    displayResults(results);
}

function displayResults(results) {
    const resultsDiv = document.getElementById('results');
    resultsDiv.innerHTML = '';

    if (results.length === 0) {
        resultsDiv.textContent = 'No results found';
        return;
    }

    // Create table to display results
    const table = document.createElement('table');
    const thead = document.createElement('thead');
    const tbody = document.createElement('tbody');

    // Create header row using the first row of excelData
    const headerRow = document.createElement('tr');
    excelData[0].forEach(header => {
        const th = document.createElement('th');
        th.textContent = header;
        headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);

    results.forEach(result => {
        const row = document.createElement('tr');
        result.forEach(cell => {
            const td = document.createElement('td');
            td.textContent = cell;
            row.appendChild(td);
        });
        tbody.appendChild(row);
    });

    table.appendChild(thead);
    table.appendChild(tbody);
    resultsDiv.appendChild(table);

    // Add "Description" and "Remarks" text below the table
    const descriptionText = document.createElement('p');
    descriptionText.textContent = 'Description';
    descriptionText.style.fontWeight = 'bold';
    resultsDiv.appendChild(descriptionText);

    const remarksText = document.createElement('p');
    remarksText.textContent = 'Remarks';
    remarksText.style.fontWeight = 'bold';
    resultsDiv.appendChild(remarksText);
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

function convertToSpreadsheet() {
  const fileInput = document.getElementById('fileInput');
  const file = fileInput.files[0];
  if (file) {
    const reader = new FileReader();
    // When the file is loaded, this function executes
    reader.onload = function (e) {
      // Read the file content as text
      const text = e.target.result;
      // Split the text into lines based on newlines
      const rows = text.trim().split('\n');

      // Create a new table element
      const table = document.createElement('table');
      
      // Arrays to store processed rows and flipped rows
      let processedRows = [];
      let flippedRows = [];
      let totalCaps = 0; // Variable to store total caps earned

      // Loop through each line of the text file
      for (let i = 0; i < rows.length; i++) {
        // Split the line into cells based on whitespace
        const cells = rows[i].trim().split(/\s+/);

        // Check if the line contains at least 3 cells
        if (cells.length >= 3) {
          // Determine if it's a caps line by checking if the second last cell is a number
          const isCapsLine = !isNaN(cells[cells.length - 2]);
          if (isCapsLine) {
            // Store caps lines in processedRows
            processedRows.push(cells);
            // Accumulate total caps
            totalCaps += parseInt(cells[cells.length - 2], 10);
          } else {
            // Store non-caps lines in flippedRows
            flippedRows.push(cells);
          }
        }
      }

      // Create header row
      const headerRow = document.createElement('tr');
      const headers = ['Character', 'Date/Time', 'Player', 'Item', 'Caps'];
      headers.forEach(headerText => {
        const header = document.createElement('th');
        header.textContent = headerText;
        headerRow.appendChild(header);
      });
      table.appendChild(headerRow);

      // Create total caps row
      const totalRow = document.createElement('tr');
      totalRow.innerHTML = `<th colspan="4"><strong>Total Caps Earned from Vendor:</strong></th><td>${totalCaps}</td>`;
      table.appendChild(totalRow);

      // Iterate through pairs of processedRows and flippedRows
      for (let i = 0; i < Math.min(processedRows.length, flippedRows.length); i++) {
        const cells1 = processedRows[i]; // Cells from the caps line
        const cells2 = flippedRows[i];   // Cells from the non-caps line

        // Extract relevant data from cells
        const character = cells1[0].trim(); // Character name
        const datetime = cells1.slice(1, 5).join(' '); // Date/time information
        const player = cells2[7] || ''; // Player name (handling undefined case)
        const item = cells2.slice(10).join(' ').trim() || ''; // Item description (handling undefined case)
        const caps = cells1[cells1.length - 2].trim(); // Caps information

        // Create a new row for the table
        const tableRow = document.createElement('tr');
        // Combine all data into an array
        const data = [character, datetime, player, item, caps];

        // Iterate through each data item and create a cell for it
        data.forEach(cellText => {
          const cell = document.createElement('td');
          cell.textContent = cellText;
          tableRow.appendChild(cell);
        });

        // Append the row to the table
        table.appendChild(tableRow);
      }

      // Append the table to an output div
      const outputDiv = document.getElementById('output');
      if (outputDiv) {
        outputDiv.innerHTML = ''; // Clear previous content
        outputDiv.appendChild(table); // Append the table
      } else {
        // If outputDiv is not found
        console.error('Output div not found');
      }

      // Convert table to CSV and download
      tocsv(table);
    };

    // Read the file as text
    reader.readAsText(file);
  } else {
    // Alert if no file is selected
    alert('Please select a file first.');
  }
}

// Function to convert table to CSV and initiate download
function tocsv(table) {
  // Convert table to array
  let data = [];
  for (let tr of table.querySelectorAll('tr')) {
    let row = [];
    for (let td of tr.querySelectorAll('td, th')) {
      row.push(td.textContent.trim());
    }
    data.push(row.join(','));
  }

  // Convert array to CSV blob
  let csvContent = data.join('\n');
  let blob = new Blob([csvContent], { type: 'text/csv' });

  // Create download link and initiate download
  let url = window.URL.createObjectURL(blob);
  let a = document.createElement('a');
  a.href = url;
  a.download = 'VendorLog.csv';
  a.click();
  a.remove();
  window.URL.revokeObjectURL(url);
}

// Global variable to hold the parsed student data from the Excel file
let studentData = [];

// Get references to the DOM elements we'll be working with
const searchForm = document.getElementById('search-form');
const studentIdInput = document.getElementById('student-id');
const courseNameInput = document.getElementById('course-name');
const searchButton = document.getElementById('search-button');
const resultsDiv = document.getElementById('results');

/**
 * Fetches and parses the exam data from the .xlsx file.
 * This function is called as soon as the script loads.
 */
async function loadExamData() {
    try {
        // Fetch the excel file from the same directory.
        // This requires running a local web server.
        const response = await fetch('exam.xlsx');
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}. Make sure 'exam.xlsx' is in the same folder and you are running a local server.`);
        }
        const arrayBuffer = await response.arrayBuffer();
        const data = new Uint8Array(arrayBuffer);

        // Use SheetJS to parse the workbook
        const workbook = XLSX.read(data, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        // Convert the sheet to a JSON object array
        // cellDates: true ensures dates are parsed as JS Date objects instead of numbers
        studentData = XLSX.utils.sheet_to_json(worksheet, { cellDates: true });

        // Populate the Course Dropdown
        const courses = [...new Set(studentData.map(item => item.course))].sort();
        courseNameInput.innerHTML = '<option value="">Select a course...</option>'; // Reset
        courses.forEach(course => {
            if (course) {
                const option = document.createElement('option');
                option.value = course;
                option.textContent = course;
                courseNameInput.appendChild(option);
            }
        });

        // Enable the search form
        searchButton.disabled = false;
        resultsDiv.textContent = 'Ready for search. Please enter your details above.';

    } catch (error) {
        console.error("Error loading or processing the Excel file:", error);
        resultsDiv.textContent = `Error: Could not load exam data. ${error.message}`;
        resultsDiv.style.color = 'red';
        searchButton.disabled = true;
    }
}

/**
 * Event listener for the search form submission.
 * This is triggered when the student clicks the "Search" button.
 */
searchForm.addEventListener('submit', (event) => {
    // Prevent the form from causing a page reload
    event.preventDefault();

    if (studentData.length === 0) {
        resultsDiv.textContent = 'Exam data is not available. Please try reloading the page.';
        return;
    }

    // Get the search values from the form, trimming whitespace
    const searchId = studentIdInput.value.trim();
    const searchCourse = courseNameInput.value.trim().toLowerCase();

    // Find the matching record in our data
    const result = studentData.find(row => {
        // Important: Compare values carefully.
        // The column names ('ID', 'course') must match your Excel file exactly.
        const recordId = String(row['ID']).trim();
        const recordCourse = String(row['course']).trim().toLowerCase();

        return recordId === searchId && recordCourse === searchCourse;
    });

    // Display the result
    displayResult(result);
});

/**
 * Renders the found result to the results div.
 * This is now dynamic to handle varying columns.
 * @param {object | undefined} result - The found student record object, or undefined if not found.
 */
function displayResult(result) {
    if (result) {
        // Define the columns that are always present and should be shown first
        const fixedColumns = ['Date', 'course', 'ID', 'name', 'Average', 'Status'];
        let resultHTML = '<h3>Result Found</h3>';

        // Display fixed columns first, in order
        fixedColumns.forEach(colName => {
            if (result[colName] !== undefined) {
                let value = result[colName];

                // Format Date to MM/YYYY
                if (colName === 'Date') {
                    // Handle Excel serial dates (numbers)
                    if (typeof value === 'number') {
                        // Convert Excel number to Date. Adding 12 hours (0.5) to avoid timezone issues at midnight.
                        value = new Date((value - 25569 + 0.5) * 86400 * 1000);
                    }

                    if (value instanceof Date) {
                        const month = String(value.getMonth() + 1).padStart(2, '0');
                        const year = value.getFullYear();
                        value = `${month}/${year}`;
                    }
                }

                // Format Average as Percentage
                if (colName === 'Average') {
                    value = (value * 100).toFixed(2) + '%';
                }

                resultHTML += `<div class="result-item"><strong>${colName}:</strong> ${value}</div>`;
            }
        });

        // Display any other columns (like Stone#1, Marks#1, etc.) dynamically
        resultHTML += '<hr style="margin: 1rem 0;"><h4>Detailed Marks:</h4>';
        for (const key in result) {
            // Check if the key is not one of the fixed columns and the object has this property
            if (Object.prototype.hasOwnProperty.call(result, key) && !fixedColumns.includes(key)) {
                let value = result[key];
                
                // Add leading zero to single digits
                if (!isNaN(value) && Number(value) < 10) {
                    value = '0' + Number(value);
                }

                // Format Marks: Theory is / 20, others are / 10
                if (key === 'Theory') {
                    value = value + ' / 20';
                } else if (key.startsWith('Stone')) {
                    // Stone columns are just the number
                } else {
                    value = value + ' / 10';
                }
                
                resultHTML += `<div class="result-item"><strong>${key}:</strong> <span style="font-family: 'Gochi Hand', cursive; font-weight: bold; font-size: 1.2rem;">${value}</span></div>`;
            }
        }
        resultsDiv.innerHTML = resultHTML;

    } else {
        resultsDiv.innerHTML = `
            <h3>No Record Found</h3>
            <p>We could not find a record matching the ID and Course you provided. Please double-check the information and try again.</p>
        `;
    }
}

// Initialise the data loading process when the page loads
loadExamData();

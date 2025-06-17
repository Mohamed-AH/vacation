// Step 1: Array of IDs to process
const idArray = [
    "60215", "60062", "60069", "60245", "60319", "60322", "60349", "60362", "60360", "60376",
    "60119", "60425", "60044", "60007", "60039", "12262", "60072", "60472", "60471", "12253",
    "12255", "60477", "60478", "60479", "60481", "10742", "60382", "60493", "60497", "60461",
    "60499", "60498", "60500", "60504", "60566", "60550", "60549", "60571", "60554", "60556",
    "60510", "60509", "60508", "60557", "60518", "60519", "60512", "60516", "60562", "60560",
    "60531", "60031", "60525", "60528", "60020", "60527", "60539", "60535", "60561", "60544",
    "60576", "60577", "60581", "60597", "60596", "60641", "60613", "60609", "60616", "60672",
    "60624", "60626", "60682", "60645", "60689", "60646", "60657", "60703", "60735", "60669",
    "60670", "60683", "60684", "60708", "60686", "60697", "60728", "60729", "60710", "60721",
    "60704", "60701", "60736", "60745", "60751", "60767", "60775", "60797", "60800", "60809"
];

// Step 2: Function to enter ID into the input field
function enterId(id) {
    const inputElement = document.getElementById("wfStatusGrid_DXSE_I");
    if (inputElement) {
        inputElement.value = id;
        inputElement.dispatchEvent(new Event('input', { bubbles: true }));
        inputElement.dispatchEvent(new Event('change', { bubbles: true }));
        console.log(`Entered ID: ${id}`);
    } else {
        console.error("Input field not found!");
    }
}

// Step 3: Function to extract table data as CSV for "Annual Leave" only
function getAnnualLeaveAsCSV() {
    const table = document.getElementById("wfStatusGrid_DXMainTable");
    if (!table) return null;

    // Extract headers
    const headers = [];
    const headerCells = table.querySelectorAll("tr[id^='wfStatusGrid_DXHeadersRow'] th, tr[id^='wfStatusGrid_DXHeadersRow'] td");

    headerCells.forEach(cell => {
        const text = cell.innerText.trim().replace(/\s+/g, ' ');
        if (text && !headers.includes(text)) {
            headers.push(text);
        }
    });

    // Remove "Details" column from headers
    const detailsIndex = headers.indexOf("");
    if (detailsIndex !== -1) headers.splice(detailsIndex, 1);

    // Extract rows for "Annual Leave"
    const rows = [];
    const dataRows = table.querySelectorAll("tr[id^='wfStatusGrid_DXDataRow']");

    dataRows.forEach(row => {
        const cells = row.querySelectorAll("td");
        const rowData = [];

        let workflowType = "";
        cells.forEach((cell, i) => {
            let text = cell.innerText.trim().replace(/\s+/g, ' ');

            // Skip "Details" button column
            if (i === 0) return;

            rowData.push(text);

            // Check if this is the Workflow Type column
            if (i === 3) { // Assuming Workflow Type is the 4th column (index 3)
                workflowType = text;
            }
        });

        // Include only rows where Workflow Type is "Annual Leave"
        if (workflowType === "Annual Leave") {
            rows.push(rowData);
        }
    });

    // Build CSV
    let csv = headers.join(",") + "\n";
    rows.forEach(row => {
        csv += row.map(cell => `"${cell.replace(/"/g, '""')}"`).join(",") + "\n";
    });

    return csv;
}

// Step 4: Main function to loop through IDs and process
async function runAutomation() {
    for (let i = 0; i < idArray.length; i++) {
        const id = idArray[i];

        console.log(`Processing ID #${i + 1}: ${id}`);

        // Enter ID into the input box
        enterId(id);

        // Wait for table to load (adjust timeout as needed)
        await new Promise(resolve => setTimeout(resolve, 3000)); // 3 seconds

        // Get Annual Leave data as CSV
        const csv = getAnnualLeaveAsCSV();
        if (csv) {
            console.log(`\n--- Annual Leave Data for ID: ${id} ---`);
            console.log(csv);
        } else {
            console.warn(`No Annual Leave data found for ID: ${id}`);
        }

        // Pause before processing the next ID
        await new Promise(resolve => setTimeout(resolve, 2000)); // 2 seconds
    }

    console.log("✅ Automation complete!");
}

// Start the automation
runAutomation();
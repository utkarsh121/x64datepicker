/*
 * TASKPANE.JS
 * The core logic for the Excel Calendar Add-in.
 * This combines date calculation and modern Excel API interaction.
 *
 * The use of the Excel.run() method ensures compatibility with
 * both 32-bit and 64-bit versions of Office.
 */

// Global variables for tracking the current month/year being displayed
let currentMonth = new Date().getMonth();
let currentYear = new Date().getFullYear();

/**
 * 1. Initialization and Setup
 * Office.onReady is the required entry point for all modern Office Add-ins.
 */
Office.onReady(async (info) => {
    if (info.host === Office.HostType.Excel) {
        // Render the calendar as soon as Office is ready
        renderCalendar(currentYear, currentMonth);

        // Hook up the navigation buttons using a single delegated listener
        document.getElementById('calendar-container').addEventListener('click', handleNavigation);
        
        // Initial status update
        document.getElementById('status-bar').innerText = "Click a date to insert it.";
    }
});


/**
 * 2. Excel API Interaction (The 64-bit compatible part)
 * This function uses the modern Excel.run API to write data.
 */
function writeDateToExcel(dateString) {
    // Check for the minimum required API set (ExcelApi 1.1)
    if (!Office.context.requirements.isSetSupported("ExcelApi", 1.1)) {
        updateStatus("Error: Excel API 1.1 required. Please update Excel.");
        return;
    }

    Excel.run(async (context) => {
        // Get the currently selected cell/range
        const selectedRange = context.workbook.getSelectedRange();
        
        // 🌟 FIX: Load the 'address' property before context.sync() 🌟
        // This prevents the "property 'address' is not available" error 
        // when trying to read it for the status update later.
        selectedRange.load("address"); 

        // Queue commands to set the value and format
        selectedRange.values = [[dateString]];
        selectedRange.numberFormat = [["m/d/yyyy"]]; 

        // Execute all queued commands (critical for all Excel.run blocks)
        await context.sync();
        
        // The address property is now available locally after sync
        updateStatus(`Inserted: ${dateString} at ${selectedRange.address}`);

    }).catch(function(error) {
        // Handle any errors during the Excel operation
        updateStatus(`Error inserting date: ${error.message}`);
    });
}


/**
 * 3. Calendar Rendering and Event Handlers
 */

// Helper to update the footer status bar
function updateStatus(message) {
    document.getElementById('status-bar').innerText = message;
}

// Generates the main calendar HTML structure for a given year and month
function renderCalendar(year, month) {
    const today = new Date();
    const date = new Date(year, month);
    const monthName = date.toLocaleString('default', { month: 'long' });
    const daysInMonth = new Date(year, month + 1, 0).getDate();
    const firstDayOfWeek = new Date(year, month, 1).getDay(); // 0 = Sunday, 1 = Monday, etc.

    let calHtml = `
        <div class="calendar-header">
            <!-- 🌟 FIX: Four distinct navigation buttons 🌟 -->
            <button class="ms-Button ms-Button--icon ms-Icon ms-Icon--DoubleChevronLeft" data-action="prev-year"></button>
            <button class="ms-Button ms-Button--icon ms-Icon ms-Icon--ChevronLeft" data-action="prev-month"></button>
            <h4 class="ms-font-l" style="display:inline-block; width: 140px;">${monthName} ${year}</h4>
            <button class="ms-Button ms-Button--icon ms-Icon ms-Icon--ChevronRight" data-action="next-month"></button>
            <button class="ms-Button ms-Button--icon ms-Icon ms-Icon--DoubleChevronRight" data-action="next-year"></button>
        </div>
        <table class="calendar-table">
            <thead>
                <tr>
                    <th>S</th><th>M</th><th>T</th><th>W</th><th>T</th><th>F</th><th>S</th>
                </tr>
            </thead>
            <tbody>
    `;
    
    // Start drawing the calendar grid
    let dayCount = 1;
    let started = false;

    for (let i = 0; i < 6; i++) { // Max 6 weeks
        calHtml += '<tr>';
        for (let j = 0; j < 7; j++) { // 7 days per week
            
            if (i === 0 && j === firstDayOfWeek) {
                // Start of the month
                started = true;
            }

            if (started && dayCount <= daysInMonth) {
                const currentDay = new Date(year, month, dayCount);
                const dateStr = currentDay.toISOString().substring(0, 0) === 'T' ? currentDay.toISOString().substring(0, 10) : currentDay.toISOString().substring(0, 10);
                
                // Check if this day is today for highlighting
                const isToday = currentDay.toDateString() === today.toDateString() ? 'today' : '';
                
                calHtml += `<td class="date-cell ${isToday}" data-date="${dateStr}">${dayCount}</td>`;
                dayCount++;
            } else {
                // Empty cells before the first day or after the last day
                calHtml += '<td></td>';
            }
        }
        calHtml += '</tr>';
        if (dayCount > daysInMonth) {
            break; // Stop loop once all days are rendered
        }
    }

    calHtml += '</tbody></table>';
    
    // Insert the generated HTML into the task pane
    document.getElementById('calendar-container').innerHTML = calHtml;

    // Attach click listeners to all generated date cells
    attachDateClickListeners();
}

// Handles clicks on the calendar (date selection)
function attachDateClickListeners() {
    document.querySelectorAll('.date-cell').forEach(cell => {
        cell.addEventListener('click', function() {
            const dateStr = this.getAttribute('data-date');
            if (dateStr) {
                writeDateToExcel(dateStr);
            }
        });
    });
}

// Handles month/year navigation
function handleNavigation(event) {
    // Check if the click target or its parent has a data-action attribute
    const action = event.target.getAttribute('data-action') || event.target.parentElement.getAttribute('data-action');
    
    if (action) {
        if (action === 'prev-month') {
            currentMonth--;
            if (currentMonth < 0) {
                currentMonth = 11;
                currentYear--;
            }
        } else if (action === 'next-month') {
            currentMonth++;
            if (currentMonth > 11) {
                currentMonth = 0;
                currentYear++;
            }
        } else if (action === 'prev-year') {
            currentYear--;
        } else if (action === 'next-year') {
            currentYear++;
        }
        
        // Re-render the calendar with the new month/year
        renderCalendar(currentYear, currentMonth);
        updateStatus("Navigated calendar.");
    }
}
const baseUrl = "https://api.bluegamma.io/v1/";

/**
 * Get swap rate
 * @customfunction
 * @param {string} token Token
 * @param {string} index Index
 * @param {string} start_date Start date
 * @param {string} maturity_date Maturity date
 * @param {string} payment_frequency Payment frequency
 * @param {string} valuation_time Valuation time
 * @returns {string} Swap rate
 */
async function SwapRate(token, index, start_date, maturity_date, payment_frequency, valuation_time = "") {
    start_date_val = await GetDate(start_date);
    maturity_date_val = await GetDate(maturity_date);

    const params = new URLSearchParams({
        index: index,
        start_date: start_date_val,
        maturity_date: maturity_date_val,
        payment_frequency: payment_frequency
    });

    if (valuation_time) {
        params.append("valuation_time", valuation_time);
    }

    const url = baseUrl + "swap_rate?" + params.toString();

    try {
        const response = await fetch(url, {
            method: 'GET',
            headers: {
                'X-Api-Key': token
            }
        });

        if (!response.ok) {
            return `HTTP error! Status: ${response.status}`;
        }

        const data = await response.json();
        console.log(data);
        return data.swap_rate;
    } catch (error) {
        console.error('Error fetching the swap rate:', error);
        // Optionally, you can return an error value or rethrow the error
        throw error;
    }
}

/**
 * Get forward rate
 * @customfunction
 * @param {string} token Token
 * @param {string} index Index
 * @param {string} start_date Start date
 * @param {string} end_date End date
 * @param {string} valuation_time Valuation time
 * @returns {string} Forward rate
 */
async function ForwardRate(token, index, start_date, end_date, valuation_time = "") {
    start_date = await GetDate(start_date);
    end_date = await GetDate(end_date);

    const params = new URLSearchParams({
        index: index,
        start_date: start_date,
        end_date: end_date
    });

    if (valuation_time) {
        params.append("valuation_time", valuation_time);
    }

    const url = baseUrl + "forward_rate?" + params.toString();

    try {
        const response = await fetch(url, {
            method: 'GET',
            headers: {
                'X-Api-Key': token
            }
        });

        if (!response.ok) {
            return `HTTP error! Status: ${response.status}`;
        }

        const data = await response.json();
        console.log(data);
        return data.forward_rate;
    } catch (error) {
        console.error('Error fetching the forward rate:', error);
        // Optionally, you can return an error value or rethrow the error
        throw error;
    }
}

async function GetDate(dateInput) {
    let date;

    // Check if dateInput is a cell reference
    if (/^[A-Z]+\d+$/.test(dateInput)) {
        // Assume dateInput is a cell reference, retrieve the value
        const range = Excel.run(async (context) => {
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const cell = sheet.getRange(dateInput);
            cell.load("values");
            await context.sync();
            return cell.values[0][0]; // Get the value of the cell
        });

        // Wait for the range operation to complete
        const cellValue = await range;

        // Try to parse the cell value as a date
        date = new Date(cellValue);
    } else {
        // Otherwise, assume it's a direct date string
        date = new Date(dateInput);
    }

    // Perform your logic with the date
    if (isNaN(date.getTime())) {
        throw new Error("Invalid date input");
    }

    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0'); // Months are zero-indexed
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
}

function formatDate(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
}

CustomFunctions.associate("SwapRate", SwapRate);
CustomFunctions.associate("ForwardRate", ForwardRate);
CustomFunctions.associate("GetDate", GetDate);
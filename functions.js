const baseUrl = "https://6v51jtul4e.execute-api.eu-west-2.amazonaws.com/v1/";
const isTokenValidName = "isTokenValid";

////////////////////////

/**
 * Get swap rate with token
 * @customfunction
 * @param {string} token Token
 * @param {string} index Index
 * @param {string} start_date Start date
 * @param {string} maturity_date Maturity date
 * @param {string} payment_frequency Payment frequency
 * @param {string} valuation_time Valuation time
 * @returns {string} Swap rate
 */
async function SwapRateWithToken(token, index, start_date, maturity_date, payment_frequency, valuation_time = "") {
    const isTokenValid = await OfficeRuntime.storage.getItem(isTokenValidName);
    if (isTokenValid != "true") {
        return;
    }

    start_date = await GetDate(start_date);
    maturity_date = await GetDate(maturity_date);

    const params = new URLSearchParams({
        index: index,
        start_date: start_date,
        maturity_date: maturity_date,
        payment_frequency: payment_frequency
    });

    if (valuation_time) {
        params.append("valuation_time", valuation_time);
    }

    const url = baseUrl + "swap_rate?" + params.toString();

    try {
        data = await GetRateWithToken(token, url);
        return data.swap_rate;
    } catch (error) {
        console.error('Error fetching the swap rate:', error);
    }
}

async function GetRateWithToken(token, url) {
    console.log("token", token);

    const response = await fetch(url, {
        method: 'GET',
        headers: {
            "Authorization": `Bearer ${token}`
        }
    });

    if (response.status === 401) {
        OfficeRuntime.storage.setItem(isTokenValidName, false);
        return `Token expired! Status: ${response.status}`;
    }
    else if (!response.ok) {
        return `HTTP error! Status: ${response.status}`;
    }

    const data = await response.json();
    console.log(data);
    return data;
}

//////////////////////////////////////

/**
 * Get swap rate
 * @customfunction
 * @param {string} index Index
 * @param {string} start_date Start date
 * @param {string} maturity_date Maturity date
 * @param {string} payment_frequency Payment frequency
 * @param {string} valuation_time Valuation time
 * @returns {string} Swap rate
 */
async function SwapRate(index, start_date, maturity_date, payment_frequency, valuation_time = "") {
    const isTokenValid = await OfficeRuntime.storage.getItem(isTokenValidName);
    if (isTokenValid != "true") {
        return;
    }

    start_date = await GetDate(start_date);
    maturity_date = await GetDate(maturity_date);

    const params = new URLSearchParams({
        index: index,
        start_date: start_date,
        maturity_date: maturity_date,
        payment_frequency: payment_frequency
    });

    if (valuation_time) {
        params.append("valuation_time", valuation_time);
    }

    const url = baseUrl + "swap_rate?" + params.toString();
    
    try {
        data = await GetRate(url);
        return data.swap_rate;
    } catch (error) {
        console.error('Error fetching the swap rate:', error);
    }
}

/**
 * Get forward rate
 * @customfunction
 * @param {string} index Index
 * @param {string} start_date Start date
 * @param {string} end_date End date
 * @param {string} valuation_time Valuation time
 * @returns {string} Forward rate
 */
async function ForwardRate(index, start_date, end_date, valuation_time = "") {
    const isTokenValid = await OfficeRuntime.storage.getItem(isTokenValidName);
    if (isTokenValid != "true") {
        return;
    }        

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
        data = await GetRate(url);
        return data.forward_rate;
    } catch (error) {
        console.error('Error fetching the forward rate:', error);
    }
}

async function GetRate(url) {
    const token = await OfficeRuntime.storage.getItem("jwtToken");
    console.log("token", token);

    const response = await fetch(url, {
        method: 'GET',
        headers: {
            "Authorization": `Bearer ${token}`
        }
    });

    if (response.status === 401) {
        OfficeRuntime.storage.setItem(isTokenValidName, false);
        return `Token expired! Status: ${response.status}`;
    }
    else if (!response.ok) {
        return `HTTP error! Status: ${response.status}`;
    }

    const data = await response.json();
    console.log(data);
    return data;
}

async function GetDate(dateInput) {
    console.log("dateInput:", dateInput);

    let date;

    // Check if dateInput is a cell reference
    if (/^-?\d+$/.test(dateInput)) {
        date = new Date((dateInput - 25569) * 86400 * 1000); // Excel epoch adjustment
        console.log("Converted Date:", date);
    } else {
        // If not a cell reference, assume it's a direct date string
        date = new Date(dateInput);
    }

    if (isNaN(date.getTime())) {
        throw new Error("Invalid date input");
    }

    // Format the date as YYYY-MM-DD
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0'); // Months are zero-indexed
    const day = String(date.getDate()).padStart(2, '0'); // Extract day
    const formattedDate = `${year}-${month}-${day}`;
    console.log("Formatted Date:", formattedDate);
    return formattedDate;
}

function formatDate(date) {
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
}

CustomFunctions.associate("SwapRate", SwapRate);
CustomFunctions.associate("ForwardRate", ForwardRate);
CustomFunctions.associate("SwapRateWithToken", SwapRateWithToken);
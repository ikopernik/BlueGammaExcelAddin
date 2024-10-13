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
async function BG_SwapRate(token, index, start_date, maturity_date, payment_frequency, valuation_time = "") {
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
async function BG_ForwardRate(token, index, start_date, end_date, valuation_time = "") {
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

CustomFunctions.associate("BG_SwapRate", BG_SwapRate);
CustomFunctions.associate("BG_ForwardRate", BG_ForwardRate);
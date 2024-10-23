/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

const WORKOS_API_KEY = "";
const WORKOS_CLIENT_ID = "client_01HRW99Z2RAQHX8R63PNJPSZ8C"; // Replace with your actual WorkOS client ID
const REDIRECT_URI = "https://ikopernik.github.io/callback.html"; // Replace with your callback URL

Office.onReady()
.then(function() {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  document.getElementById("run").onclick = run;       
});

async function run() {
  try {
      const authorizationUrl = await getAuthorizationUrl();
  } catch (error) {
    console.error(error);
  }
}


async function getAuthorizationUrl() {
    const redirectUri = encodeURIComponent('https://ikopernik.github.io/BlueGammaExcelAddin/callback.html');
    const endpoint = `https://dev.bluegamma.io/auth/url?redirectUri=${redirectUri}`;

    try {
        //const response = await fetch(endpoint, { mode: 'no-cors' });
        const response = await fetch(endpoint);
        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }
        const data = await response.json();
        return data.url; // Assuming the response contains an `url` field
    } catch (error) {
        console.error('Error fetching authorization URL:', error);
    }
}

function getClientId() {
    if (!WORKOS_CLIENT_ID) {
        throw new Error("WORKOS_CLIENT_ID is not set");
    }
    return WORKOS_CLIENT_ID;
}
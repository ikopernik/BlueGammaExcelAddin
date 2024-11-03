let isInitialized = false;

Office.onReady().then(async function () {
    // Check if the initialization has already been done
    if (isInitialized) {
        return; // If already initialized, exit the function
    }
    isInitialized = true;

    // Ensure the DOM is loaded before setting up the button click handler
    document.getElementById("loginButton").addEventListener("click", authenticateUser);

    console.log("Office is ready.");
    // Check if the user is already authenticated on load
    await checkAuthenticationStatus();
});

function authenticateUser() {
    // URL for the authentication endpoint
    const authUri = "https://dev.bluegamma.io/api/auth/addin?redirectUri=https://ikopernik.github.io/BlueGammaExcelAddin/callback.html";

    // window.location.href = authUri;
    window.open(authUri, "_blank");

    /////////////

    //if (Office.context.requirements.isSetSupported("ExcelApi", "1.3")) {
    //    window.location.href = authUri;
    //} else {
    //    window.open(authUri, "_blank");
    //}

    //Office.context.ui.messageParent("Some message", { targetOrigin: authUri });

    /////////////

    // Add an event listener for messages from the child window (callback.html)
    window.addEventListener("message", async function (event) {
        // Check the origin of the message for security
        if (event.origin === "https://ikopernik.github.io") { // Replace with your actual domain
            if (event.data.type === "AUTH_SUCCESS") {
                console.log("authorizationCode", event.data.authorizationCode);

                try {
                    // Fetch the JWT token using the authentication code
                    const response = await fetch(`https://dev.bluegamma.io/api/auth/jwt?code=${encodeURIComponent(event.data.authorizationCode)}`);
                    if (!response.ok) {
                        throw new Error(`HTTP error! status: ${response.status}`);
                    }

                    const data = await response.json();
                    const jwtToken = data.token;
                    console.log("Received JWT Token:", jwtToken);

                    // Store the JWT token for future use
                    OfficeRuntime.storage.setItem("jwtToken", jwtToken);
                    OfficeRuntime.storage.setItem(isTokenValidName, true);
                    let authenticationStarted = false;

                    // Update UI to show authenticated status
                    await UpdateControlsToAuthenticated();
                } catch (error) {
                    console.error("Error fetching JWT token:", error);
                    document.getElementById("authStatus").textContent = "Error fetching JWT token";
                }

            } else if (event.data.type === "AUTH_FAILURE") {
                console.log("Authentication failed.");
                document.getElementById("authStatus").textContent = "Authentication failed";
            }
        }
    });
}

async function checkAuthenticationStatus() {
    // Retrieve the JWT token from Office Roaming Settings
    const jwtToken = await OfficeRuntime.storage.getItem("jwtToken");
    const isTokenValid = await OfficeRuntime.storage.getItem(isTokenValidName);
    console.log("checkAuthenticationStatus token", jwtToken);

    if (jwtToken && isTokenValid == "true") {
        await UpdateControlsToAuthenticated();
    } else {
        await UpdateControlsToNotAuthenticated();
    }
}

let authenticationStarted = false;

async function shouldAuthenticateAgain()
{
    if (!isInitialized || authenticationStarted) {
        return;
    }

    const isTokenValid = await OfficeRuntime.storage.getItem(isTokenValidName);
    if (isTokenValid != "true") {
        authenticationStarted = true;
        await UpdateControlsToNotAuthenticated();
        Office.addin.showAsTaskpane()
    }
}

async function UpdateControlsToAuthenticated() {
    document.getElementById("authStatus").textContent = "Authenticated";
    document.getElementById("loginButton").style.display = "none";
}

async function UpdateControlsToNotAuthenticated() {
    document.getElementById("authStatus").textContent = "Not authenticated";
    document.getElementById("loginButton").style.display = "block";
}

// Periodically check for updates (e.g., every second)
setInterval(shouldAuthenticateAgain, 1000);
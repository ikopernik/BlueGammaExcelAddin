let isInitialized = false;
const isTokenValidName = "isTokenValid";

document.addEventListener("DOMContentLoaded", function () {
    // Check if the initialization has already been done
    if (isInitialized) {
        return; // If already initialized, exit the function
    }
    isInitialized = true;

    // Ensure the DOM is loaded before setting up the button click handler
    document.getElementById("loginButton").addEventListener("click", authenticateUser);

    console.log("Office is ready.");
    const jwtToken = OfficeRuntime.storage.getItem("jwtToken");
    // Check if the user is already authenticated on load
    checkAuthenticationStatus();
});

function authenticateUser() {
    // URL for the authentication endpoint
    const authUri = "https://dev.bluegamma.io/api/auth/addin?redirectUri=https://ikopernik.github.io/BlueGammaExcelAddin/callback.html";

    // window.location.href = authUri;
    window.open(authUri, "_blank");

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
                    document.getElementById("authStatus").textContent = "Authenticated";
                    document.getElementById("loginButton").style.display = "none";
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

function checkAuthenticationStatus() {
    // Retrieve the JWT token from Office Roaming Settings
    const jwtToken = OfficeRuntime.storage.getItem("jwtToken");
    const isTokenValid = OfficeRuntime.storage.setItem(isTokenValidName);
    console.log("checkAuthenticationStatus token", jwtToken);

    if (jwtToken && isTokenValid) {
        document.getElementById("authStatus").textContent = "Authenticated";
        document.getElementById("loginButton").style.display = "none";
    } else {
        document.getElementById("authStatus").textContent = "Not authenticated";
        document.getElementById("loginButton").style.display = "block";
    }
}

let authenticationStarted = false;

async function shouldAuthenticateAgain()
{
    const isTokenValid = await OfficeRuntime.storage.getItem(isTokenValidName);
    if (!isTokenValid && !authenticationStarted) {
        authenticationStarted = true;

        // Your logic to show or bring attention to the task pane
        document.getElementById("taskPaneContent").style.display = "block";

        document.getElementById("authStatus").textContent = "Not authenticated";
        document.getElementById("loginButton").style.display = "block";
    }
}

// Periodically check for updates (e.g., every second)
setInterval(checkForTaskPaneFlag, 1000);

// Initialize SecureLS
//const ls = new SecureLS({ encodingType: 'aes' });

// Storing the JWT token
//function storeToken(token) {
//    ls.set('jwtToken', token);
//}

// Retrieving the JWT token
//function getToken() {
//    return ls.get('jwtToken');
//}
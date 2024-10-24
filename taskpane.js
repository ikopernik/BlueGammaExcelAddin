Office.onReady().then(function() {
    // Ensure the DOM is loaded before setting up the button click handler
    document.getElementById("loginButton").addEventListener("click", authenticateUser);

    console.log("Office is ready.");
    const jwtToken = localStorage.getItem("jwtToken");
    // Check if the user is already authenticated on load
    checkAuthenticationStatus();
});

function authenticateUser() {
    // URL for the authentication endpoint
    const authUri = "https://dev.bluegamma.io/api/auth/addin?redirectUri=https://ikopernik.github.io/BlueGammaExcelAddin/callback.html";

    // Use the Office Dialog API to open the authentication page
    Office.context.ui.displayDialogAsync(authUri, { height: 60, width: 30 }, (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
            console.error("Failed to open the dialog:", result.error.message);
        } else {
            const dialog = result.value;

            // Add an event listener for messages from the child window
            window.addEventListener("message", function (event) {
                // Check the origin of the message for security
                if (event.origin === "https://ikopernik.github.io") { // Replace with your actual domain
                    if (event.data.type === "AUTH_SUCCESS") {
                        console.log("authorizationCode", event.data.authorizationCode);

                        // Save the token
                        //localStorage.setItem("jwtToken", event.data.authorizationCode);

                        // Update UI to show authenticated status
                        document.getElementById("authStatus").textContent = "Authenticated";
                        document.getElementById("loginButton").style.display = "none";
                    } else if (event.data.type === "AUTH_FAILURE") {
                        console.log("Authentication failed.");
                        document.getElementById("authStatus").textContent = "Authentication failed";
                    }
                }
            });

            // Handle the dialog being closed
            dialog.addEventHandler(Office.EventType.DialogEventReceived, () => {
                console.log("Dialog was closed by the user");
            });
        }
    });
}

function checkAuthenticationStatus() {
    // Retrieve the JWT token from Office Roaming Settings
    const jwtToken = localStorage.getItem("jwtToken");

    if (jwtToken) {
        document.getElementById("authStatus").textContent = "Authenticated";
        document.getElementById("loginButton").style.display = "none";
    } else {
        document.getElementById("authStatus").textContent = "Not authenticated";
        document.getElementById("loginButton").style.display = "block";
    }
}
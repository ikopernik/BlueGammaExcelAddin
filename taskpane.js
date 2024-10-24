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

            // Handle messages sent from the dialog
            dialog.addEventHandler(Office.EventType.DialogMessageReceived, (messageResult) => {
                console.log("Message in taskpane received");
                const jwtToken = messageResult.message;
                dialog.close();

                if (jwtToken) {
                    // Save the token
                    localStorage.setItem("jwtToken", jwtToken);

                    // Update UI to show authenticated status
                    document.getElementById("authStatus").textContent = "Authenticated";
                    document.getElementById("loginButton").style.display = "none";
                } else {
                    document.getElementById("authStatus").textContent = "Authentication failed";
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
window.onload = function () {
    console.log("Callback is called");

    const jwtToken = getCookie("token");

    console.log(jwtToken);

    if (jwtToken) {
        // If the token is found, send it to the parent page (the task pane)
        if (window.opener) {
            window.opener.postMessage({ type: "AUTH_SUCCESS", token: jwtToken }, "*");
        }
    } else {
        // If the token is not found, send an empty message to indicate failure
        if (window.opener) {
            window.opener.postMessage({ type: "AUTH_FAILURE" }, "*");
        }
    }

    // Optionally, close the child window after sending the message
    // window.close();
};

// Function to get the JWT token from cookies
function getCookie(name) {
    const value = `; ${document.cookie}`;
    console.log(value);
    const parts = value.split(`; ${name}=`);
    if (parts.length === 2) return parts.pop().split(";").shift();
    return null;
}
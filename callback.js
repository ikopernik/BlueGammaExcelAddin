window.onload = function () {
    console.log("Callback is called");
    console.log(window.location);
    console.log(document.cookie);

    const params = new URLSearchParams(window.location.search);
    console.log(params);

    const authorizationCode = params.get('code');
    console.log("Authorization Code:", authorizationCode);

    if (authorizationCode) {
        // If the token is found, send it to the parent page (the task pane)
        if (window.opener) {
            window.opener.postMessage({ type: "AUTH_SUCCESS", authorizationCode: authorizationCode }, "*");
        }
    } else {
        // If the token is not found, send an empty message to indicate failure
        if (window.opener) {
            window.opener.postMessage({ type: "AUTH_FAILURE" }, "*");
        }
    }

    // Optionally, close the child window after sending the message
    window.close();
};
window.onload = function () {
    console.log("Callback is called");
    // Function to get the JWT token from cookies
    function getCookie(name) {
        const value = `; ${document.cookie}`;
        const parts = value.split(`; ${name}=`);
        if (parts.length === 2) return parts.pop().split(";").shift();
        return null;
    }

    const jwtToken = getCookie("token");

    if (jwtToken) {
        // Send the JWT token to the parent page (the task pane)
        Office.context.ui.messageParent(jwtToken);
    } else {
        // Send an empty message to indicate failure
        Office.context.ui.messageParent("");
    }
};
// Extract the authorization code from the URL
async function handleAuthCallback() {
    const urlParams = new URLSearchParams(window.location.search);
    const authCode = urlParams.get('code');

    if (!authCode) {
        console.error('Authorization code not found in the callback URL.');
        alert('Failed to authenticate. Please try again.');
        return;
    }

    try {
        // Exchange the authorization code for a JWT token
        const tokenResponse = await fetch(`https://dev.bluegamma.io/auth/user-jwt?code=${encodeURIComponent(authCode)}`, {
            method: 'GET'
        });

        if (!tokenResponse.ok) {
            throw new Error(`Failed to get JWT token. HTTP status: ${tokenResponse.status}`);
        }

        const tokenData = await tokenResponse.json();

        // Assuming the response contains a 'token' field
        const jwtToken = tokenData.token;

        if (!jwtToken) {
            throw new Error('JWT token not found in the response.');
        }

        // Save the token to Office's roaming settings so it's available across sessions
        Office.onReady(() => {
            Office.context.roamingSettings.set('jwtToken', jwtToken);
            Office.context.roamingSettings.saveAsync((result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    console.log('JWT token saved successfully.');
                    alert('Authentication successful! You may close this window.');
                    // Optionally, redirect to another page or perform further actions
                } else {
                    console.error('Failed to save JWT token:', result.error.message);
                    alert('Failed to save authentication token. Please try again.');
                }
            });
        });

    } catch (error) {
        console.error('Error handling authentication callback:', error);
        alert('An error occurred during authentication. Please try again.');
    }
}

// Call the function to handle the authentication callback
handleAuthCallback();
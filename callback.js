import { SignJWT } from 'jose'; // Import the SignJWT function from the jose library
import { WorkOS } from '@workos-inc/node'; // Make sure WorkOS SDK is available

export const workos = new WorkOS(process.env.WORKOS_API_KEY);

export function getClientId() {
    const clientId = process.env.WORKOS_CLIENT_ID;

    if (!clientId) {
        throw new Error("WORKOS_CLIENT_ID is not set");
    }

    return clientId;
}

export function getJwtSecretKey() {
    const secretKey = process.env.JWT_SECRET_KEY;

    if (!secretKey) {
        throw new Error("JWT_SECRET_KEY is not set");
    }

    return secretKey;
}

export function generateJWT(payload) {
    return new SignJWT(payload)
        .setProtectedHeader({ alg: "HS256", typ: "JWT" })
        .setIssuedAt()
        .setExpirationTime("30d")
        .sign(getJwtSecretKey());
}

// Function to set the token when it is retrieved
function setJwtToken(token) {
    jwtToken = token;
}

// Save the token in roaming settings and set the global variable
function saveAndSetToken(token) {
    Office.context.roamingSettings.set("jwtToken", token);
    Office.context.roamingSettings.saveAsync(function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            console.log("JWT token saved successfully.");
            setJwtToken(token);
        } else {
            console.error("Failed to save JWT token:", asyncResult.error.message);
        }

        // Send the token back to the parent task pane to complete the authentication process
        Office.context.ui.messageParent(JSON.stringify({ token: firstToken }));
    });
}

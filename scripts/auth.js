const msalParams = {
    auth: {
        authority: "https://login.microsoftonline.com/consumers",
        clientId: "853384a0-f07f-4a6b-92d0-54bef089c7b4",
        redirectUri: location.toString()
    },
}

const app = new msal.PublicClientApplication(msalParams);

async function getToken() {

    let accessToken = "";

    authParams = { scopes: ["OneDrive.ReadOnly"] };

    try {

        // see if we have already the idtoken saved
        const resp = await app.acquireTokenSilent(authParams);
        accessToken = resp.accessToken;

    } catch (e) {

        // per examples we fall back to popup
        const resp = await app.loginPopup(authParams);
        app.setActiveAccount(resp.account);

        if (resp.idToken) {

            const resp2 = await app.acquireTokenSilent(authParams);
            accessToken = resp2.accessToken;

        }
    }

    return accessToken;
}

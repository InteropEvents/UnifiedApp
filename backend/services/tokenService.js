require('dotenv').config({ path: require('path').resolve(__dirname, '../.env') });

const msal = require('@azure/msal-node');

const config = {
    auth: {
        clientId: process.env.CLIENT_ID,
        authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
        clientSecret: process.env.CLIENT_SECRET,
    }
};

const cca = new msal.ConfidentialClientApplication(config);

async function getAppLevelToken() {
    const tokenRequest = {
        scopes: [process.env.SCOPE || "https://graph.microsoft.com/.default"],
    };

    try {
        const response = await cca.acquireTokenByClientCredential(tokenRequest);
        return response.accessToken;
    } catch (error) {
        console.error("Error acquiring token:", error);
        throw error;
    }
}

module.exports = { getAppLevelToken };

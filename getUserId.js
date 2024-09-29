import { Client } from '@microsoft/microsoft-graph-client';
import 'isomorphic-fetch';
import { ClientSecretCredential } from '@azure/identity';
import dotenv from 'dotenv';
dotenv.config();

// Example usage
const email = 'alexei.80@hotmail.fr';

const clientId = process.env.CLIENT_ID;
const tenantId = process.env.TENANT_ID;
const clientSecret = process.env.CLIENT_SECRET;


getUserIdByEmail(tenantId, clientId, clientSecret, email)
    .then(userId => console.log(`User ID is: ${userId}`))
    .catch(error => console.error(error));

async function getUserIdByEmail(tenantId, clientId, clientSecret, email) {

    const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);

    // Use the credential to get an access token for the Microsoft Graph API
    const tokenResponse = await credential.getToken('https://graph.microsoft.com/.default');
    const accessToken = tokenResponse.token;

    // Initialize Microsoft Graph Client with the token
    const client = Client.init({
        authProvider: (done) => {
            done(null, accessToken); // Provide the access token
        },
    });

    try {
        // Call Microsoft Graph API to get user by email (userPrincipalName)
        const user = await client.api(`/users/${email}`).get();

        // Return the user ID
        console.log(`User ID: ${user.id}`);
        return user.id;
    } catch (error) {
        console.error('Error getting user by email:', error);
        throw error;
    }
}

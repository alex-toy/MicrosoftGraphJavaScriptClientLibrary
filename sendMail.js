// import 'isomorphic-fetch';
import { Client } from '@microsoft/microsoft-graph-client';
import { ClientSecretCredential } from '@azure/identity';
import dotenv from 'dotenv';
dotenv.config();



export async function sendEmail(message) {
    const client = await _getClient();

    // try {
    //     await client.api(`/users/${userId}/sendMail`).post(message);
    //     console.log('Email sent successfully!');
    // } catch (error) {
    //     console.error('Error sending email', error)
    // }

    console.log(message);
}

export async function createOnlineMeeting(onlineMeeting){
    const client = await _getClient();
    const userId = process.env.USER_ID;
  
    await client.api(`/users/${userId}/onlineMeetings`).post(onlineMeeting);
}

async function _getClient() {
    const clientId = process.env.CLIENT_ID;
    const tenantId = process.env.TENANT_ID;
    const clientSecret = process.env.CLIENT_SECRET;

    const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);

    const client = Client.init({
        authProvider: async (done) => {
            try {
                const tokenResponse = await credential.getToken('https://graph.microsoft.com/.default');
                done(null, tokenResponse.token);
            } catch (error) {
                done(error, null);
            }
        },
    });
    
    return client;
}

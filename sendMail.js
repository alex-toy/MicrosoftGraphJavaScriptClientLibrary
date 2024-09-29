import { Client } from '@microsoft/microsoft-graph-client';
import { ClientSecretCredential } from '@azure/identity';
import dotenv from 'dotenv';
dotenv.config();

sendEmail();

async function sendEmail() {
    const clientId = process.env.CLIENT_ID;
    const tenantId = process.env.TENANT_ID;
    const clientSecret = process.env.CLIENT_SECRET;
    const userId = process.env.USER_ID;

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

    const message = {
        message: {
            subject: 'Test Email from Microsoft Graph',
            body: {
                contentType: 'Text',
                content: 'Hello from Microsoft Graph API!',
            },
            toRecipients: [
                {
                    emailAddress: {
                        address: 'tekumsee.80@gmail.com',
                    },
                },
            ],
        },
        saveToSentItems: true,
    };

    try {
        await client.api(`/users/${userId}/sendMail`).post(message);
        console.log('Email sent successfully!');
    } catch (error) {
        console.error('Error sending email', error)
    }
}
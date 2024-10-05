// import 'isomorphic-fetch';
// import { Client } from '@microsoft/microsoft-graph-client';
// import { ClientSecretCredential } from '@azure/identity';
import dotenv from 'dotenv';
dotenv.config();



export async function sendEmail(message) {
    const client = await _getClient();
    const userId = process.env.USER_ID;

    // const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);

    // const client = Client.init({
    //     authProvider: async (done) => {
    //         try {
    //             const tokenResponse = await credential.getToken('https://graph.microsoft.com/.default');
    //             done(null, tokenResponse.token);
    //         } catch (error) {
    //             done(error, null);
    //         }
    //     },
    // });

    // try {
    //     await client.api(`/users/${userId}/sendMail`).post(message);
    //     console.log('Email sent successfully!');
    // } catch (error) {
    //     console.error('Error sending email', error)
    // }

    console.log(message);
    console.log(clientId)
}

async function createOnlineMeeting(){
    const client = await _getClient();
    const userId = process.env.USER_ID;

    const onlineMeeting = {
        startDateTime: '2024-09-30T14:00:00.2444915-07:00',
        endDateTime: '2024-09-30T15:00:00.2464912-07:00',
        subject: 'Backlog Refinement',
        participants: {
            organizer: {
                upn: "formateur@sparks-formation.com",
                role: "presenter"
            },
            attendees: [
                {
                    upn: "alessio.rea@apollossc.com",
                    role: "attendee"
                }
            ]
        }
    };
  
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

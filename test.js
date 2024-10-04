// Import the TokenCredential class that you wish to use. This example uses a Client SecretCredential
import 'isomorphic-fetch';
import { ClientSecretCredential } from "@azure/identity";
import { Client } from "@microsoft/microsoft-graph-client";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js";

// formateur graphtest
const clientId = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx";
const tenantId = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx";
const clientSecret = "";
const userId = "xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx";

await sendEmail();
// await createOnlineMeeting();
// await getUser();
// await createEvent();

async function getUser(){
    let _client = await _getClient();
    const res = await await _client.api(`/users/${userId}`).get();
    console.log(res)
}

async function createOnlineMeeting(){
    let _client = await _getClient();

    const onlineMeeting = {
        startDateTime: '2019-07-25T14:00:00.2444915-07:00',
        endDateTime: '2019-07-25T15:00:00.2464912-07:00',
        subject: 'Backlog Refinement',
        participants: {
            organizer: {
                upn: "test@sparks.com",
                role: "presenter"
            },
            attendees: [
                {
                    upn: "alessio.rea@test.com",
                    role: "attendee"
                }
            ]
        }
    };
  
    await _client.api(`/users/${userId}/onlineMeetings`).post(onlineMeeting);
}

async function createEvent(){
    let _client = await _getClient();

    const event = {
        subject: "Ma réunion",
        start: {
            dateTime: "2024-09-25T12:09:07.149Z",
            timeZone: "UTC"
        },
        end: {
            dateTime: "2024-09-25T12:10:07.149Z",
            timeZone: "UTC"
        },
        body : {
            contentType : "html",
            content : "<h1>join url</h1>"
        },
        attendees: [
            {
                emailAddress: {
                    address: "alexei.80@hotmail.fr",
                    name: "alexei"
                },
                type: "required"
            }
        ]
    };
    
    await _client.api(`/users/${userId}/events`).post(event);
}

async function _getClient(){
    const tokenCredential = new ClientSecretCredential(tenantId, clientId, clientSecret);
    
    const authProvider = new TokenCredentialAuthenticationProvider(tokenCredential, { scopes: ['https://graph.microsoft.com/.default'] });
    
    let _client = Client.initWithMiddleware({ authProvider: authProvider });

    return _client;
}

async function sendEmail() {
    const client = await _getClient();

    const message = {
        message: {
            subject: 'Mail refacto test.js',
            body: {
                contentType: 'Text',
                content: "Hello, si vous recevez ce mail, c'est que je viens de réussir à utiliser l'API de microsoft graph!!!!!",
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

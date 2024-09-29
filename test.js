import "isomorphic-fetch"; // or import the fetch polyfill you installed
import { Client } from "@microsoft/microsoft-graph-client";
import { ClientCertificateCredential } from "@azure/identity";
import { TokenCredentialAuthenticationProvider } from "@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js";

const credential = new ClientSecretCredential(
    'YOUR_TENANT_ID',
    'YOUR_CLIENT_ID',
    'YOUR_CLIENT_SECRET',
  );
  
const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ['https://graph.microsoft.com/.default'],
});

const graphClient = Client.initWithMiddleware({ authProvider: authProvider });

let userDetails = await graphClient.api("/me").get();
console.log(userDetails);
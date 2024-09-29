# Microsoft Graph JavaScript Client Library

```
npm i isomorphic-fetch
npm install @microsoft/microsoft-graph-client
npm install @azure/msal-browser
npm install @azure/identity
```

- create a .env file

- add your sensitive variables to it, like this:
```
USER_ID=your-user-id-here
```

- npm install dotenv
```
npm install dotenv
```

- in the js file, you can retrieve those variables :
```
import dotenv from 'dotenv';
dotenv.config();
const userId = process.env.USER_ID;
```


https://github.com/microsoftgraph/msgraph-sdk-javascript/blob/dev/docs/TokenCredentialAuthenticationProvider.md
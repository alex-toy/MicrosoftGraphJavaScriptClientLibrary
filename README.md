# Microsoft Graph JavaScript Client Library

## Node Server

- create project
```
npm init
npm i express
```

- create docker image
```
docker build -t graphapi_i:v1 .
```

- create container
```
docker run --name graphapi_c -p 8080:8080 -v ${pwd}:/app -v /node_modules graphapi_i:v1
```


## Secret Management

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

## Add Packages

```
npm i isomorphic-fetch
npm install @microsoft/microsoft-graph-client
npm install @azure/msal-browser
npm install @azure/identity
```


https://github.com/microsoftgraph/msgraph-sdk-javascript/blob/dev/docs/TokenCredentialAuthenticationProvider.md
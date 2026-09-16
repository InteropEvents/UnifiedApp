# Getting Started with Create React App

This sample demonstrates the integration of the Microsoft Graph Toolkit into a fully functional React application using [Create React App](https://github.com/facebook/create-react-app).

The sample uses the library [@microsoft/mgt-react](https://www.npmjs.com/package/@microsoft/mgt-react) to simplify usage of [Microsoft Graph Toolkit (mgt)](https://aka.ms/mgt) web components in React. The library wraps all mgt components and exports them as React components.

## Prerequisites
1. Microsoft 365 Developer Account
2. Node.js and yarn installed
3. Register your APP and get your own Application (Client ID). For more details, see section **Registering Application with Entra ID** below
4. Configure API permissions
    - After registering, go to the "API permissions" section in your App registration
    - Click "Add a permission" and select "Microsoft Graph"
    - Choose "Delegated permissions"
    - Choose the permissions that application needs (e.g., "User.Read", "Mail.Read", etc.). Check InteropSolutions\UnifiedApp\frontend\src\index.tsx for list of APIs in scope.
    - Click "Add permissions"
5. Configure Authentication settings
    - Open Authentication section of your app registration
    - Scroll down to the "Implicit grant and hybrid flows" section
    - Enable the checkbox for "Access tokens (used for implicit flows)". Make sure that "ID tokens" box is disabled

## Registering Application with Entra ID
1. Open https://entra.microsoft.com/#home
2. Applications/App registrations
3. Click "+ New registration"
4. Enter Name (ex.: "UnifiedApp"), choose "Accounts in any organizational directory (Any Microsoft Entra ID tenant - Multitenant) and personal Microsoft accounts (e.g. Skype, Xbox)"
5. Redirect URI: Select "Single-page application (SPA)"; enter http://localhost as Redirect URI
6. Click Register
7. Save Application (client) ID
  
## Setting up this sample

1. Clone this demo samples from github repo
2. run `cd InteropSolutions/UnifiedApp/frontend` to app folder
3. Copy the `.env.sample` to its own `.env` file
4. Add your own Application (Client ID) in the file for `REACT_APP_CLIENT_ID`
5. Add your own Open AI API_KEY in the file for `REACT_APP_API_KEY`
6. Install the dependencies with `yarn`

## Available Scripts

In the project directory, you can run:

### `yarn start`

Runs the app in the development mode.<br>
Open [http://localhost:3000](http://localhost:3000) to view it in the browser.

The page will reload if you make edits.<br>
You will also see any lint errors in the console.

### `yarn build`

Builds the app for production to the `build` folder.<br>
It correctly bundles React in production mode and optimizes the build for the best performance.

The build is minified and the filenames include the hashes.<br>
Your app is ready to be deployed!

## Deployment to Azure only for front end ( no ci/cd )
1. Create a static Wep App
2. Copy deployment token.
3. use following command to deploy:

C:\repos\CaaS\InteropSolutions\UnifiedApp\frontend> swa deploy $build_folder_location --deployment-token $Deployment_token --env test 

$build_folder_location = C:\repos\CaaS\InteropSolutions\UnifiedApp\frontend\build <br>
$Deployment_token = Get it from Portal. <br>
-- env = test or prod etc NOTE environment name will decide the url of the site. To get to the correct url  <br>
Go to azure portal -> your static web app -> settings -> environment -> preview environment


See the section about [deployment](https://facebook.github.io/create-react-app/docs/deployment) for more information.



## Learn More

You can learn more in the [Create React App documentation](https://facebook.github.io/create-react-app/docs/getting-started).

To learn React, check out the [React documentation](https://reactjs.org/).

### Code Splitting

This section has moved here: https://facebook.github.io/create-react-app/docs/code-splitting

### Analyzing the Bundle Size

This section has moved here: https://facebook.github.io/create-react-app/docs/analyzing-the-bundle-size

### Making a Progressive Web App

This section has moved here: https://facebook.github.io/create-react-app/docs/making-a-progressive-web-app

### Advanced Configuration

This section has moved here: https://facebook.github.io/create-react-app/docs/advanced-configuration

### Deployment

This section has moved here: https://facebook.github.io/create-react-app/docs/deployment

### `npm run build` fails to minify

This section has moved here: https://facebook.github.io/create-react-app/docs/troubleshooting#npm-run-build-fails-to-minify

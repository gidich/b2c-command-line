import { setupEnvironment } from './setup-environment.js';
setupEnvironment();
import readline, { ReadLine } from 'readline';
import chalk from 'chalk';

const rl = readline.createInterface({
    input: process.stdin,
    output: process.stdout
});

interface TokenResponse {
    token_type: string;
    expires_in: number;
    ext_expires_in: number;
    access_token: string;
}

// https://learn.microsoft.com/en-us/graph/auth-v2-service?context=graph%2Fapi%2F1.0&view=graph-rest-1.0&tabs=http#4-request-an-access-token
const requestToken  =  async  (tenant:string, clientId:string, clientSecret:string): Promise<TokenResponse> => {
    const result = await fetch(`https://login.microsoftonline.com/${tenant}/oauth2/v2.0/token`, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/x-www-form-urlencoded'
      },
        body:  new URLSearchParams({
          client_id: clientId,
          scope: 'https://graph.microsoft.com/.default',
          client_secret: clientSecret,
          grant_type: 'client_credentials'
      }),
        redirect: 'follow'
      });
    return result.json();
}

// https://learn.microsoft.com/en-us/graph/auth-v2-service?context=graph%2Fapi%2F1.0&view=graph-rest-1.0&tabs=curl#5-use-the-access-token-to-call-microsoft-graph
//https://learn.microsoft.com/en-us/graph/extensibility-overview?tabs=javascript#retrieve-a-directory-extension
const listUsers = async (token: TokenResponse) => {
    const result = await fetch('https://graph.microsoft.com/beta/users', {
        method: 'GET',
        headers: {
            'Authorization': `${token.token_type} ${token.access_token}`
        }
    });
    return result.json();
}

interface UserSpecification {
    accountEnabled?: boolean;
    displayName: string;
    userPrincipalName: string;
    passwordProfile: {  
        password: string;
        forceChangePasswordNextSignIn: boolean;
    }
    identities: [{
        signInType: string;
        issuer: string;
        issuerAssignedId: string;
    }]
}

const addExtensionAttributeToUser = (userSpec: UserSpecification, extensionAppId:string, extensionAttribute:string, extensionAttributeValue:string): UserSpecification => {
    const extensionName:string =`extension_${extensionAppId.replaceAll('-','')}_${extensionAttribute}`;
    (userSpec as any)[extensionName] = extensionAttributeValue;
    return userSpec;
}

const createUserSpecification = (displayName: string,  password: string, emailAddress:string) => {
    const userSpecification: UserSpecification = {
        accountEnabled: true,
        displayName: displayName,
        userPrincipalName: emailAddress.replace('@', '_')+'#EXT#@'+`${process.env.TENANT_NAME}.onmicrosoft.com`,
        passwordProfile: {
            password: password,
            forceChangePasswordNextSignIn: false,
        },
        identities: [{
            signInType: 'emailAddress',
            issuer: `${process.env.TENANT_NAME}.onmicrosoft.com`,
            issuerAssignedId: emailAddress
        }]
    }
    return userSpecification;
}

const addUser = async (token: TokenResponse, userSpecification: UserSpecification) => {
    const result = await fetch('https://graph.microsoft.com/v1.0/users', {
        method: 'POST',
        headers: {
            'Authorization': `${token.token_type} ${token.access_token}`,
            'Content-Type': 'application/json'
        },
        body: JSON.stringify(userSpecification)
    });
    return result.json();
}

const addApplicant = async(token:TokenResponse) => {
    if(!process.env.EXTENSION_APP_ID) throw new Error('EXTENSION_APP_ID not set')
    const applicantName = await askQuestion(' Name: ');
    const applicantEmail = await askQuestion(' Email: ');
    const applicantPassword = await askQuestion('Applicant Password: ');
    const applicant = createUserSpecification(applicantName, applicantPassword, applicantEmail);
    let applicantWithExtension = addExtensionAttributeToUser(applicant, process.env.EXTENSION_APP_ID, 'accountType', 'Applicant');
    applicantWithExtension = addExtensionAttributeToUser(applicantWithExtension, process.env.EXTENSION_APP_ID, 'migrationRequired', 'true');
    console.log(chalk.blue(JSON.stringify(applicantWithExtension)));
    const result = await addUser(token, applicantWithExtension);
    console.log(chalk.green(JSON.stringify(result)));
}   

const addEntity = async(token:TokenResponse) => {
    if(!process.env.EXTENSION_APP_ID) throw new Error('EXTENSION_APP_ID not set')
    const name = await askQuestion(' Name: ');
    const email = await askQuestion(' Email: ');
    const entityName = await askQuestion('Entity Name: ');
    const entityId = await askQuestion('Entity Id: ');
    const password = await askQuestion('Password: ');
    const entity = createUserSpecification(name, password, email);
    let entityWithExtension = addExtensionAttributeToUser(entity, process.env.EXTENSION_APP_ID, 'accountType', 'Entity');
    entityWithExtension = addExtensionAttributeToUser(entityWithExtension, process.env.EXTENSION_APP_ID, 'migrationRequired', 'true');
    entityWithExtension = addExtensionAttributeToUser(entityWithExtension, process.env.EXTENSION_APP_ID, 'entityName', entityName);
    entityWithExtension = addExtensionAttributeToUser(entityWithExtension, process.env.EXTENSION_APP_ID, 'entityId', entityId);
    console.log(chalk.blue(JSON.stringify(entityWithExtension)));
    const result = await addUser(token, entityWithExtension);
    console.log(chalk.green(JSON.stringify(result)));
}

const askQuestion = async (question: string) => {
    return new Promise<string>((resolve, reject) => {
        rl.question(question, (answer) => {
            resolve(answer);
        });
    });
}

const handleInput = async (input: string, rl:ReadLine, token:TokenResponse) => {
    switch (input.trim()) {
        case 'help':
            console.log(chalk.green('list - list all entities'));
            console.log(chalk.green('add-applicant - add an applicant'));
            console.log(chalk.green('add-entity - add an entity'));
            console.log(chalk.green('quit - quit the program'));
            break;
        case 'list':
            console.log(chalk.green('list'));
            const users = await listUsers(token);
            console.log(chalk.green(JSON.stringify(users)));
            break;
        case 'add-applicant':
            console.log(chalk.green('add-applicant'));
            await addApplicant(token);
            break;
        case 'add-entity':
            console.log(chalk.green('add-entity'));
            await addEntity(token);
            break;
        case 'quit':
            console.log(chalk.green('quit'));
            rl.close();
            process.exit(0);
        default:
            console.log(chalk.red('Invalid command'));
            break;
    }
}

const main = async () => {
    console.clear();
    console.log(chalk.green('Welcome to the Entity Manager'));
    console.log(chalk.green('Type "help" for a list of commands'));

    if(!process.env.TENANT_ID) throw new Error('TENANT_ID not set')
    if(!process.env.CLIENT_ID) throw new Error('CLIENT_ID not set')
    if(!process.env.CLIENT_SECRET) throw new Error('CLIENT_SECRET not set')
    
    const token = await requestToken(process.env.TENANT_ID, process.env.CLIENT_ID, process.env.CLIENT_SECRET);

    rl.setPrompt('ready> ');
    rl.prompt();
    rl.on('line', async (input) => {
        await handleInput(input, rl, token);
        rl.prompt();
    });
}

main();
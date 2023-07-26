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
const listUsers = async (token: TokenResponse) => {
    const result = await fetch('https://graph.microsoft.com/v1.0/users', {
        method: 'GET',
        headers: {
            'Authorization': `${token.token_type} ${token.access_token}`
        }
    });
    return result.json();
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
            break;
        case 'add-entity':
            console.log(chalk.green('add-entity'));
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
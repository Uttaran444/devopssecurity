import { PublicClientApplication, LogLevel, Configuration } from '@azure/msal-node';

// Azure DevOps AAD application ID used for tokens
// Using the resource's default scope pattern for Azure DevOps
const AZDO_SCOPE = ['499b84ac-1321-427f-aa17-267ca6975798/.default'];

// Buffer in seconds to refresh slightly before expiry
const EXPIRY_BUFFER = 60;

export class AuthManager {
    private msal: PublicClientApplication;
    private tokenCache: { accessToken: string; expiresAt: number } | null = null;

    constructor() {
        const tenant = process.env.TENANT_ID || 'common';
        const clientId = process.env.CLIENT_ID || '04f0c124-f73c-4b74-9a29-7ede364e3eb3'; // Microsoft first-party public client fallback

        const config: Configuration = {
            auth: {
                clientId,
                authority: `https://login.microsoftonline.com/${tenant}`,
            },
            system: {
                loggerOptions: { loggerCallback: () => {}, logLevel: LogLevel.Warning, piiLoggingEnabled: false },
            },
        };
        this.msal = new PublicClientApplication(config);
    }

    public async getAuthToken(): Promise<string> {
        if (this.isTokenValid()) {
            return this.tokenCache!.accessToken;
        }
        return this.deviceCodeLogin();
    }

    private isTokenValid(): boolean {
        return !!this.tokenCache && this.tokenCache.expiresAt > Date.now();
    }

    // Interactive device-code login for end users
    private async deviceCodeLogin(): Promise<string> {
        const result = await this.msal.acquireTokenByDeviceCode({
            scopes: AZDO_SCOPE,
            deviceCodeCallback: (response) => {
                console.log(`To sign in, visit ${response.verificationUri} and enter code: ${response.userCode}`);
                console.log(response.message);
            },
        });
        if (!result || !result.accessToken) {
            throw new Error('Device code login did not return an access token');
        }
        const expiresOnMs = result.expiresOn ? result.expiresOn.getTime() : Date.now() + 3600 * 1000;
        const expiresAt = expiresOnMs - EXPIRY_BUFFER * 1000;
        this.tokenCache = { accessToken: result.accessToken, expiresAt };
        return result.accessToken;
    }
}
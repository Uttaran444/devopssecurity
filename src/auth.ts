// A small buffer in seconds to ensure we refresh the token before it actually expires.
const EXPIRY_BUFFER = 60;

/**
 * Manages the authentication token for Dynamics 365, including fetching and caching.
 */
export class AuthManager {
    private tokenCache: {
        accessToken: string;
        expiresAt: number;
    } | null = null;

    /**
     * Retrieves a valid access token, refreshing if necessary.
     * @returns {Promise<string>} A valid bearer token.
     */
    public async getAuthToken(): Promise<string> {
        if (this.isTokenValid()) {
            console.log('Using cached auth token.');
            return this.tokenCache!.accessToken;
        }

        console.log('Auth token is invalid or expired. Fetching a new one...');
        return this.fetchNewToken();
    }

    /**
     * Checks if the cached token is still valid.
     * @returns {boolean} True if the token is valid, false otherwise.
     */
    private isTokenValid(): boolean {
        if (!this.tokenCache) {
            return false;
        }
        return this.tokenCache.expiresAt > Date.now();
    }

    /**
     * Fetches a new OAuth token from Azure AD.
     * @returns {Promise<string>} The new access token.
     */
    private async fetchNewToken(): Promise<string> {
        const { TENANT_ID, CLIENT_ID, CLIENT_SECRET, DYNAMICS_RESOURCE_URL } = process.env;

        if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET || !DYNAMICS_RESOURCE_URL) {
            throw new Error('Missing required environment variables for authentication in .env file.');
        }

        const tokenUrl = `https://login.microsoftonline.com/${TENANT_ID}/oauth2/token`;

        const params = new URLSearchParams();
        params.append('grant_type', 'client_credentials');
        params.append('client_id', CLIENT_ID);
        params.append('client_secret', CLIENT_SECRET);
        params.append('resource', DYNAMICS_RESOURCE_URL);

        try {
            const response = await fetch(tokenUrl, {
                method: 'POST',
                body: params,
                headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
            });

            if (!response.ok) {
                const errorText = await response.text();
                throw new Error(`Failed to fetch auth token: ${response.status} ${errorText}`);
            }

            const data = await response.json();
            const expiresInSeconds = parseInt(data.expires_in, 10);
            const expiresAt = Date.now() + (expiresInSeconds - EXPIRY_BUFFER) * 1000;

            this.tokenCache = { accessToken: data.access_token, expiresAt };

            console.log('Successfully fetched and cached new auth token.');
            return this.tokenCache.accessToken;
        } catch (error) {
            console.error('Error during token fetch:', error);
            throw error;
        }
    }
}
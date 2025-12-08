import { PublicClientApplication, DeviceCodeRequest, AuthenticationResult } from '@azure/msal-node';
import 'dotenv/config';

//const clientId = process.env.AZDO_AAD_CLIENT_ID || '9bb671b6-b073-432c-958e-9ee2d67412f3';
const clientId = '9bb671b6-b073-432c-958e-9ee2d67412f3';
//const tenantId = process.env.AZDO_AAD_TENANT_ID || '27a4ed1d-793a-4844-99db-0021f00e4a97';
const tenantId = '27a4ed1d-793a-4844-99db-0021f00e4a97';
const authority = `https://login.microsoftonline.com/${tenantId}`;

const scopes = (process.env.AZDO_AAD_SCOPES || '499b84ac-1321-427f-aa17-267ca6975798/.default')
  .split(',')
  .map(s => s.trim())
  .filter(Boolean);

let pca: PublicClientApplication | null = null;
let cachedToken: { accessToken: string; expiresOn?: number } | null = null;

function ensureClient(): PublicClientApplication {
  if (!pca) {
    if (!clientId) {
      throw new Error('AZDO_AAD_CLIENT_ID is not set.');
    }
    pca = new PublicClientApplication({ auth: { clientId, authority } });
  }
  return pca;
}

function isTokenValid(): boolean {
  if (!cachedToken || !cachedToken.accessToken) return false;
  if (!cachedToken.expiresOn) return true;
  return Date.now() < (cachedToken.expiresOn - 60 * 1000);
}

export async function getAccessToken(sendNotification?: (n: any) => Promise<void> | void): Promise<string> {
  if (isTokenValid()) return cachedToken!.accessToken;

  const client = ensureClient();
  const request: DeviceCodeRequest = {
    scopes,
    deviceCodeCallback: (response) => {
      if (sendNotification) {
        sendNotification({
          method: 'notifications/message',
          params: { level: 'info', data: `Device Code Login: ${response.message}` }
        });
      } else {
        console.log('[DeviceCode]', response.message);
      }
    }
  };

  let result: AuthenticationResult | null = null;
  try {
    result = await client.acquireTokenByDeviceCode(request);
  } catch (e: any) {
    if (sendNotification) {
      await sendNotification({ method: 'notifications/message', params: { level: 'error', data: `Device code auth failed: ${e.message || String(e)}` } });
    }
    throw e;
  }

  if (!result || !result.accessToken) {
    throw new Error('Failed to obtain access token via device code flow.');
  }

  cachedToken = {
    accessToken: result.accessToken,
    expiresOn: result.expiresOn ? result.expiresOn.getTime() : undefined
  };

  if (sendNotification) {
    await sendNotification({ method: 'notifications/message', params: { level: 'info', data: 'User login successful. Token acquired.' } });
  }

  return result.accessToken;
}

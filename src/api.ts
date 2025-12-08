import { CallToolResult } from '@modelcontextprotocol/sdk/types.js';
import { AuthManager } from './auth.js';
import { URL } from 'url';

const authManager = new AuthManager();

// Helper function to safely send notifications
async function safeNotify(sendNotification: (notification: any) => void | Promise<void>, notification: any): Promise<void> {
    try {
        await sendNotification(notification);
    } catch (error) {
        // Silently ignore notification errors
    }
}

export async function makeApiCall(
    method: 'GET' | 'POST' | 'PATCH',
    url: string,
    body: Record<string, unknown> | null,
    sendNotification: (notification: any) => void | Promise<void>
): Promise<CallToolResult> {
    try {
        await safeNotify(sendNotification, {
            method: "notifications/message",
            params: { level: "info", data: `Calling ${method} ${url}` }
        });

        const token = await authManager.getAuthToken();

        const response = await fetch(url, {
            method: method,
            headers: {
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json',
                'Accept': 'application/json, text/xml',
                // PAGINATION: Tell OData we prefer paginated responses
                'Prefer': 'odata.maxpagesize=100'
            },
            ...(body && { body: JSON.stringify(body) }),
        });

        if (response.status === 204) {
            return { content: [{ type: 'text', text: 'Operation successful (No Content).' }] };
        }

        const responseText = await response.text();

        if (!response.ok) {
            await safeNotify(sendNotification, {
                method: "notifications/message",
                params: { level: "error", data: `API call failed with status ${response.status}: ${responseText}` }
            });
            try {
                const errorJson = JSON.parse(responseText);
                const prettyError = JSON.stringify(errorJson, null, 2);
                return { isError: true, content: [{ type: 'text', text: `API Error: ${response.status}\n${prettyError}` }] };
            } catch (e) {
                return { isError: true, content: [{ type: 'text', text: `API Error: ${response.status}\n${responseText}` }] };
            }
        }

        const contentType = response.headers.get("Content-Type");
        if (contentType?.includes("text/plain") || contentType?.includes("application/xml")) {
            return { content: [{ type: 'text', text: responseText }] };
        }

        try {
            const jsonResponse = JSON.parse(responseText);
            
            // PAGINATION: Check for the next link and construct a helpful message.
            const nextLink = jsonResponse['@odata.nextLink'];
            let resultText = JSON.stringify(jsonResponse, null, 2);

            if (nextLink) {
                const nextUrl = new URL(nextLink);
                const skipParam = nextUrl.searchParams.get('$skip');
                const topParam = nextUrl.searchParams.get('$top');
                
                const paginationHint = `\n\n---
[INFO] More data is available. To get the next page, call the 'odataQuery' tool again with the parameter: "skip": ${skipParam}.`;

                resultText += paginationHint;

                await safeNotify(sendNotification, {
                    method: "notifications/message",
                    params: { level: "info", data: `More data available. Next skip token is ${skipParam}.` }
                });
            }

            return { content: [{ type: 'text', text: resultText }] };
        } catch {
            return { content: [{ type: 'text', text: responseText }] };
        }

    } catch (error) {
        const errorMessage = error instanceof Error ? error.message : String(error);
        console.error(`Error in makeApiCall: ${errorMessage}`);
        await safeNotify(sendNotification, {
            method: "notifications/message",
            params: { level: "error", data: `An unexpected error occurred: ${errorMessage}` }
        });
        return { isError: true, content: [{ type: 'text', text: `An unexpected error occurred: ${errorMessage}` }] };
    }
}
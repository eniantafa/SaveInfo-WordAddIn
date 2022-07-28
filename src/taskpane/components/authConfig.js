/* eslint-disable prettier/prettier */

import { LogLevel } from "@azure/msal-browser";


export const msalConfig = {
    auth: {
        clientId: "9613eac2-ab85-4b76-97dd-2abb746d1086",
        authority: "https://login.microsoftonline.com/common",
    redirectUri: "http://localhost:3000/",
    },
    cache: {
        cacheLocation: "sessionStorage", // This configures where your cache will be stored
        storeAuthStateInCookie: false, // Set this to "true" if you are having issues on IE11 or Edge
    },
    system: {
        loggerOptions: {
            loggerCallback: (level, message, containsPii) => {
                if (containsPii) {
                    return;
                }
                switch (level) {
                    case LogLevel.Error:
                        // eslint-disable-next-line no-undef
                        console.error(message);
                        return;
                    case LogLevel.Info:
                        // eslint-disable-next-line no-undef
                        console.info(message);
                        return;
                    case LogLevel.Verbose:
                        // eslint-disable-next-line no-undef
                        console.debug(message);
                        return;
                    case LogLevel.Warning:
                        // eslint-disable-next-line no-undef
                        console.warn(message);
                        return;
                }
            }
        }
    }
};

/**
 * Scopes you add here will be prompted for user consent during sign-in.
 * By default, MSAL.js will add OIDC scopes (openid, profile, email) to any login request.
 * For more information about OIDC scopes, visit: 
 * https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-permissions-and-consent#openid-connect-scopes
 */
export const loginRequest = {
    scopes: ["User.Read"]
};

/**
 * Add here the scopes to request when obtaining an access token for MS Graph API. For more information, see:
 * https://github.com/AzureAD/microsoft-authentication-library-for-js/blob/dev/lib/msal-browser/docs/resources-and-scopes.md
 */
export const graphConfig = {
    graphMeEndpoint: "https://graph.microsoft.com/v1.0/me"
};

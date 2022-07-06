/* eslint-disable prettier/prettier */
import { graphConfig } from "./authConfig";

/**
 * Attaches a given access token to a MS Graph API call. Returns information about the user
 * @param accessToken 
 */
export async function callMsGraph(accessToken) {
    const headers = new Headers();
    const bearer = `Bearer ${accessToken}`;

    headers.append("Authorization", bearer);

    const options = {
        method: "GET",
        headers: headers
    };

    // eslint-disable-next-line no-undef
    return fetch(graphConfig.graphMeEndpoint, options)
        .then(response => response.json())
        // eslint-disable-next-line no-undef
        .catch(error => console.log(error));
}

import * as g from '@battis/gas-lighter';
import { PREFIX } from '../Constants';

export enum ResponseFormat {
    JSON,
    Array,
    Raw,
}

let lastResponse = null;

export function getService() {
    return OAuth2.createService(`${PREFIX}.SKY`)
        .setAuthorizationBaseUrl('https://oauth2.sky.blackbaud.com/authorization')
        .setTokenUrl('https://oauth2.sky.blackbaud.com/token')
        .setClientId(g.PropertiesService.getScriptProperty('SKY_CLIENT_ID'))
        .setClientSecret(g.PropertiesService.getScriptProperty('SKY_CLIENT_SECRET'))
        .setPropertyStore(PropertiesService.getUserProperties())
        .setCache(CacheService.getUserCache())
        .setLock(LockService.getUserLock())
        .setCallbackFunction(AuthCallback);
}

export function resetService() {
    getService().reset();
}

export function makeRequest(
    url: string,
    {
        method = 'get',
        headers = {},
        ...fetchParams
    }: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {}
) {
    const service = getService();
    var maybeAuthorized = service.hasAccess();
    if (maybeAuthorized) {
        headers['Bb-Api-Subscription-Key'] =
            g.PropertiesService.getScriptProperty('SKY_ACCESS_KEY');
        headers['Authorization'] = `Bearer ${service.getAccessToken()}`;
        lastResponse = UrlFetchApp.fetch(url, {
            method,
            headers,
            ...fetchParams,
            muteHttpExceptions: true,
        });
        const code = lastResponse.getResponseCode();
        if (code >= 200 && code < 300) {
            return JSON.parse(lastResponse.getContentText('utf-8'));
        } else if (code == 401 || code == 403) {
            maybeAuthorized = false;
        } else {
            console.error(
                `API server error (${code}): ${lastResponse.getContentText('utf-8')}`
            );
            throw `API server error ${code}`;
        }
    }

    if (!maybeAuthorized) {
        CardService.newAuthorizationException()
            .setCustomUiCallback(AuthorizationCard)
            .throwException();
    }
}

export function getLastResponse() {
    return lastResponse;
}

export function authorizationCard() {
    return [
        g.CardService.newCard({
            header: 'Authorization Required',
            widgets: [
                'This add-on needs access to your Blackbaud account. You need to give permission to this add-on to make calls to the Blackbaud SKY API on your behalf.',
                CardService.newTextButton()
                    .setText('Authorize')
                    .setAuthorizationAction(
                        CardService.newAuthorizationAction().setAuthorizationUrl(
                            getService().getAuthorizationUrl()
                        )
                    ),
            ],
        }),
    ];
}

export function authCallback(request) {
    if (getService().handleCallback(request)) {
        return HtmlService.createHtmlOutput('Success! You can close this tab.');
    } else {
        return HtmlService.createHtmlOutput('Denied. You can close this tab.');
    }
}

global.action_sky_authorizationCard = authorizationCard;
const AuthorizationCard = 'action_sky_authorizationCard';

global.action_sky_authCallback = authCallback;
const AuthCallback = 'action_sky_authCallback';

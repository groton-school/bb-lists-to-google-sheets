import { Terse } from '@battis/google-apps-script-helpers';
import { PREFIX } from '../Constants';

export default class ServiceManager {
    public static getService() {
        return OAuth2.createService(`${PREFIX}.SKY`)
            .setAuthorizationBaseUrl('https://oauth2.sky.blackbaud.com/authorization')
            .setTokenUrl('https://oauth2.sky.blackbaud.com/token')
            .setClientId(Terse.PropertiesService.getScriptProperty('SKY_CLIENT_ID'))
            .setClientSecret(
                Terse.PropertiesService.getScriptProperty('SKY_CLIENT_SECRET')
            )
            .setPropertyStore(PropertiesService.getUserProperties())
            .setCache(CacheService.getUserCache())
            .setLock(LockService.getUserLock())
            .setCallbackFunction(AuthCallback);
    }

    public static resetService() {
        ServiceManager.getService().reset();
    }

    public static makeRequest(
        url: string,
        {
            method = 'get',
            headers = {},
            ...fetchParams
        }: GoogleAppsScript.URL_Fetch.URLFetchRequestOptions = {}
    ) {
        const service = ServiceManager.getService();
        var maybeAuthorized = service.hasAccess();
        if (maybeAuthorized) {
            headers['Bb-Api-Subscription-Key'] =
                Terse.PropertiesService.getScriptProperty('SKY_ACCESS_KEY');
            headers['Authorization'] = `Bearer ${service.getAccessToken()}`;
            const response = UrlFetchApp.fetch(url, {
                method,
                headers,
                ...fetchParams,
                muteHttpExceptions: true,
            });
            const code = response.getResponseCode();
            if (code >= 200 && code < 300) {
                return JSON.parse(response.getContentText('utf-8'));
            } else if (code == 401 || code == 403) {
                maybeAuthorized = false;
            } else {
                console.error(
                    `API server error (${code}): ${response.getContentText('utf-8')}`
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

    public static authorizationCard() {
        return [
            CardService.newCardBuilder()
                .setHeader(Terse.CardService.newCardHeader('Authorization Required'))
                .addSection(
                    CardService.newCardSection()
                        .addWidget(
                            Terse.CardService.newTextParagraph(
                                'This add-on needs access to your Blackbaud account. You need to give permission to this add-on to make calls to the Blackbaud SKY API on your behalf.'
                            )
                        )
                        .addWidget(
                            CardService.newTextButton()
                                .setText('Authorize')
                                .setAuthorizationAction(
                                    CardService.newAuthorizationAction().setAuthorizationUrl(
                                        ServiceManager.getService().getAuthorizationUrl()
                                    )
                                )
                        )
                )
                .build(),
        ];
    }

    public static authCallback(request) {
        if (ServiceManager.getService().handleCallback(request)) {
            return HtmlService.createHtmlOutput('Success! You can close this tab.');
        } else {
            return HtmlService.createHtmlOutput('Denied. You can close this tab.');
        }
    }
}

global.action_sky_authorizationCard = ServiceManager.authorizationCard;
const AuthorizationCard = 'action_sky_authorizationCard';
global.action_sky_authCallback = ServiceManager.authCallback;
const AuthCallback = 'action_sky_authCallback';

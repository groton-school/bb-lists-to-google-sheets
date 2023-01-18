import { Terse } from '@battis/google-apps-script-helpers';
import { LOGO_URL } from './Constants';

type SkyResponseData<T> = {
    count: number;
    value: T[];
};

export type List = {
    id: number;
    name: string;
    type: 'Basic' | 'Advanced';
    description: string;
    category: string;
    created_by: string;
    created: string;
    last_modified: string;
};

type ColumnData = {
    name: string;
    value: string;
};
type RowData = {
    columns: ColumnData[];
};
type ListData = {
    count: number;
    page: number;
    results: {
        rows: RowData[];
    };
};

export enum SkyResponse {
    Raw = 'raw',
    JSON = 'json',
    Array = 'array',
}

export default class SKY {
    public static readonly RESOURCE_NAME = 'Blackbaud SKY API';
    public static readonly URL_AUTH =
        'https://oauth2.sky.blackbaud.com/authorization';
    public static readonly URL_TOKEN = 'https://oauth2.sky.blackbaud.com/token';
    public static readonly HEADER_ACCESS_KEY: 'Bb-Api-Subscription-Key';
    public static readonly PROP_CLIENT = 'SKY_CLIENT_ID';
    public static readonly PROP_SECRET = 'SKY_CLIENT_SECRET';
    public static readonly PROP_ACCESS_KEY = 'SKY_ACCESS_KEY';

    public static school = class {
        public static v1 = class {
            public static lists(
                list_id = null,
                format = SkyResponse.JSON,
                page = 1
            ): any {
                if (list_id) {
                    const response = SKY.call(
                        `https://api.sky.blackbaud.com/school/v1/lists/advanced/${list_id}?page=${page}`
                    ) as ListData;
                    switch (format) {
                        case SkyResponse.JSON:
                            if (response.count == 0) {
                                return [];
                            }
                            return response.results.rows.map((row) => {
                                const obj = {};
                                for (const column of row.columns) {
                                    obj[column.name] = column.value;
                                }
                                return obj;
                            });
                        case SkyResponse.Array:
                            if (response.count == 0) {
                                return [];
                            }
                            const array = response.results.rows.map((row) => {
                                const arr = [];
                                for (const column of row.columns) {
                                    arr.push(column.value);
                                }
                                return arr;
                            });
                            array.unshift(
                                response.results.rows[0].columns.map(({ name, value }) => name)
                            );
                            return array;
                        case SkyResponse.Raw:
                        default:
                            return response;
                    }
                } else {
                    const response = SKY.call(
                        `https://api.sky.blackbaud.com/school/v1/lists`
                    ) as SkyResponseData<List>;
                    switch (format) {
                        case SkyResponse.JSON:
                            return response.value;
                        case SkyResponse.Raw:
                        default:
                            return response;
                    }
                }
            }
        };
    };

    private static service: GoogleAppsScriptOAuth2.OAuth2Service = null;

    public static getService() {
        if (!SKY.service) {
            const scriptProperties = PropertiesService.getScriptProperties();
            SKY.service = OAuth2.createService(SKY.RESOURCE_NAME)
                .setAuthorizationBaseUrl(SKY.URL_AUTH)
                .setTokenUrl(SKY.URL_TOKEN)
                .setClientId(scriptProperties.getProperty(SKY.PROP_CLIENT))
                .setClientSecret(scriptProperties.getProperty(SKY.PROP_SECRET))
                .setCallbackFunction('action_SKY_callbackAuthorization')
                .setPropertyStore(PropertiesService.getUserProperties())
                .setCache(CacheService.getUserCache())
                .setLock(LockService.getUserLock());
        }
        return SKY.service;
    }

    public static call(
        url: string,
        method: GoogleAppsScript.URL_Fetch.HttpMethod = 'get',
        headers = {}
    ) {
        const service = SKY.getService();
        const scriptProperties = PropertiesService.getScriptProperties();
        var maybeAuthorized = service.hasAccess();
        if (maybeAuthorized) {
            const accessToken = service.getAccessToken();
            headers[SKY.HEADER_ACCESS_KEY] = scriptProperties.getProperty(
                SKY.PROP_ACCESS_KEY
            );
            headers['Authorization'] = Utilities.formatString(
                'Bearer %s',
                accessToken
            );
            const response = UrlFetchApp.fetch(url, {
                method,
                headers,
                muteHttpExceptions: true,
            });
            const code = response.getResponseCode();
            if (code >= 200 && code < 300) {
                return JSON.parse(response.getContentText('utf-8'));
            } else if (code == 401 || code == 403) {
                maybeAuthorized = false;
            } else {
                console.error(
                    'Backend server error (%s): %s',
                    code.toString(),
                    response.getContentText('utf-8')
                );
                throw 'Backend server error: ' + code;
            }
        }

        if (!maybeAuthorized) {
            CardService.newAuthorizationException()
                .setCustomUiCallback('action_SKY_cardAuthorization')
                .throwException();
        }
    }

    public static resetService() {
        SKY.getService().reset();
    }

    public static cardAuthorization() {
        return [
            CardService.newCardBuilder()
                .setHeader(Terse.CardService.newCardHeader('Authorization Required'))
                .addSection(
                    CardService.newCardSection()
                        .addWidget(
                            CardService.newGrid().addItem(
                                CardService.newGridItem().setImage(
                                    // FIXME logo is not loading
                                    CardService.newImageComponent().setImageUrl(LOGO_URL)
                                )
                            )
                        )
                        .addWidget(
                            Terse.CardService.newTextParagraph(
                                'This add-on needs access to your Blackbaud account. You need to give permission to this add-on to make calls to the Blackbaud SKY API on your behalf.'
                            )
                        )
                        .addWidget(
                            CardService.newButtonSet().addButton(
                                CardService.newTextButton()
                                    .setText('Begin Authorization')
                                    .setAuthorizationAction(
                                        CardService.newAuthorizationAction().setAuthorizationUrl(
                                            SKY.getService().getAuthorizationUrl()
                                        )
                                    )
                            )
                        )
                )
                .build(),
        ];
    }

    public static callbackAuthorization(callbackRequest) {
        const authorized = SKY.getService().handleCallback(callbackRequest);
        // TODO better UI on API authorization
        if (authorized) {
            return HtmlService.createHtmlOutput('<h1>Authorized</h1>');
        } else {
            return HtmlService.createHtmlOutput('<h1>Denied</h1>');
        }
    }
}

global.action_SKY_callbackAuthorization = SKY.callbackAuthorization;
global.action_SKY_cardAuthorization = SKY.cardAuthorization;

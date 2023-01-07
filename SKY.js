/** global: App, Drive, Intent, Lists, Sheets, SKY, State, TerseCardService, DriveApp, SpreadsheetApp, CardService, HtmlService, PropertiesService, CacheService, LockService, OAuth2, UrlFetchApp */

const SKY = {
  RESOURCE_NAME: "Blackbaud SKY API",
  URL_AUTH: "https://oauth2.sky.blackbaud.com/authorization",
  URL_TOKEN: "https://oauth2.sky.blackbaud.com/token",
  HEADER_ACCESS_KEY: "Bb-Api-Subscription-Key",
  PROP_CLIENT: "SKY_CLIENT_ID",
  PROP_SECRET: "SKY_CLIENT_SECRET",
  PROP_ACCESS_KEY: "SKY_ACCESS_KEY",

  Response: {
    Raw: Symbol("raw"),
    JSON: Symbol("json"),
    Array: Symbol("array"),
  },

  school: {
    v1: {
      lists: (list_id = null, format = SKY.Response.JSON, page = 1) => {
        if (list_id) {
          const response = SKY.call(
            `https://api.sky.blackbaud.com/school/v1/lists/advanced/${list_id}?page=${page}`
          );
          switch (format) {
            case SKY.Response.JSON:
              if (response.count == 0) {
                return [];
              }
              return response.results.rows.map((row) => {
                const obj = {};
                for ({ name, value } of row.columns) {
                  obj[name] = value;
                }
                return obj;
              });
            case SKY.Response.Array:
              if (response.count == 0) {
                return [];
              }
              const array = response.results.rows.map((row) => {
                const arr = [];
                for ({ name, value } of row.columns) {
                  arr.push(value);
                }
                return arr;
              });
              array.unshift(
                response.results.rows[0].columns.map(({ name, value }) => name)
              );
              return array;
            case SKY.Response.Raw:
            default:
              return response;
          }
        } else {
          const response = SKY.call(
            "https://api.sky.blackbaud.com/school/v1/lists"
          );
          switch (format) {
            case SKY.Response.JSON:
              return response.value;
            case SKY.Response.Raw:
            default:
              return response;
          }
        }
      },
    },
  },

  _service_: null,

  getService: () => {
    if (!SKY._service_) {
      const scriptProperties = PropertiesService.getScriptProperties();
      SKY._service_ = OAuth2.createService(SKY.RESOURCE_NAME)
        .setAuthorizationBaseUrl(SKY.URL_AUTH)
        .setTokenUrl(SKY.URL_TOKEN)
        .setClientId(scriptProperties.getProperty(SKY.PROP_CLIENT))
        .setClientSecret(scriptProperties.getProperty(SKY.PROP_SECRET))
        .setCallbackFunction("__Sky_callbackAuthorization")
        .setPropertyStore(PropertiesService.getUserProperties())
        .setCache(CacheService.getUserCache())
        .setLock(LockService.getUserLock());
    }
    return SKY._service_;
  },

  call: (url, method = "get", headers = {}) => {
    const service = SKY.getService();
    const scriptProperties = PropertiesService.getScriptProperties();
    var maybeAuthorized = service.hasAccess();
    if (maybeAuthorized) {
      const accessToken = service.getAccessToken();
      headers[SKY.HEADER_ACCESS_KEY] = scriptProperties.getProperty(
        SKY.PROP_ACCESS_KEY
      );
      headers["Authorization"] = Utilities.formatString(
        "Bearer %s",
        accessToken
      );
      const response = UrlFetchApp.fetch(url, {
        method,
        headers,
        muteHttpExceptions: true,
      });
      const code = response.getResponseCode();
      if (code >= 200 && code < 300) {
        return JSON.parse(response.getContentText("utf-8"));
      } else if (code == 401 || code == 403) {
        maybeAuthorized = false;
      } else {
        console.error(
          "Backend server error (%s): %s",
          code.toString(),
          resp.getContentText("utf-8")
        );
        throw "Backend server error: " + code;
      }
    }

    if (!maybeAuthorized) {
      CardService.newAuthorizationException()
        .setCustomUiCallback("__Sky.cardAuthorization")
        .throwException();
    }
  },

  resetService: () => {
    SKY.getService().reset();
  },

  cardAuthorization: () => {
    return [
      CardService.newCardBuilder()
        .setHeader(TerseCardService.newCardHeader("Authorization Required"))
        .addSection(
          CardService.newCardSection()
            .addWidget(
              CardService.newGrid().addItem(
                CardService.newGridItem().setImage(
                  CardService.newImageComponent().setImageUrl(App.LOGO_URL)
                )
              )
            )
            .addWidget(
              TerseCardService.newTextParagraph(
                "This add-on needs access to your Blackbaud account. You need to give permission to this add-on to make calls to the Blackbaud SKY API on your behalf."
              )
            )
            .addWidget(
              CardService.newButtonSet().addButton(
                CardService.newTextButton()
                  .setText("Begin Authorization")
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
  },

  callbackAuthorization: (callbackRequest) => {
    const authorized = SKY.getService().handleCallback(callbackRequest);
    if (authorized) {
      return HtmlService.createHtmlOutput("<h1>Authorized</h1>");
    } else {
      return HtmlService.createHtmlOutput("<h1>Denied</h1>");
    }
  },
};

function __Sky_cardAuthorization() {
  return SKY.cardAuthorization();
}

function __Sky_callbackAuthorization(callbackRequest) {
  return SKY.callbackAuthorization(callbackRequest);
}

# Blackbaud-to-Google Lists

Google Workspace Add-on to export/sync Blackbaud lists to Google Sheets

## Install

[Refer to my gist on Google Workspace Add-on publication](https://gist.github.com/battis/6e32031196316acd1b5e5700b328aef6#file-readme-md) for general directions on how to administratively publish a Google Workspace Add-on.

This add-on is a `Sheets Add-on` in its GCP app configuration.

## Blackbaud Setup

Requires `Platform Manager` role on Blackbaud to access the list of lists
from the SKY API.

Blackbaud credentials are stored as Script Properties:

| Script Property     | Source                                                                                       |
| ------------------- | -------------------------------------------------------------------------------------------- |
| `SKY_ACCESS_KEY`    | A Blackbaud [subscription access key](https://developer.blackbaud.com/subscriptions/)        |
| `SKY_CLIENT_ID`     | The OAuth Client ID of a [Blackbaud App](https://developer.blackbaud.com/apps/)              |
| `SKY_CLIENT_SECRET` | The OAuth Client Secret of (the same) [Blackbaud App](https://developer.blackbaud.com/apps/) |

Add the script as a Redirect URI on that Blackbaud app with a URL in the pattern:

```
https://script.google.com/macros/d/:script_id/usercallback
```

## Privacy Policy

This app stores no data separately from the Google Sheets that are managed. The metadata stored in the Google Sheets describes the Blackbaud list from which to refresh the data in the Sheet.

## Terms of Service

No warranty or liability is given or implied. Use at your own risk.

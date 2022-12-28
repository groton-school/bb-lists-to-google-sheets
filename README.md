# Bb Lists to Google Sheets

Google Workspace Add-on to export/sync Blackbaud lists to Google Sheets

# Setup

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

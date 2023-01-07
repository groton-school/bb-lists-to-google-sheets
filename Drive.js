/** global: App, Drive, Intent, Lists, Sheets, SKY, State, TerseCardService, DriveApp, SpreadsheetApp, CardService, HtmlService, PropertiesService, CacheService, LockService, OAuth2, UrlFetchApp */

const Drive = {
  inferFolder: (launchEvent) => {
    if (
      launchEvent &&
      launchEvent.drive &&
      launchEvent.drive &&
      launchEvent.drive.activeCursorItem
    ) {
      const file = DriveApp.getFileById(launchEvent.drive.activeCursorItem.id);
      const parents = file.getParents();
      return parents.next() || null;
    }
    return null;
  },
};

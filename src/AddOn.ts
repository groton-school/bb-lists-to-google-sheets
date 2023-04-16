import * as WorkFlow from './Workflow';

export function onOpen() {
  SpreadsheetApp.getUi()
    .createAddonMenu()
    .addItem('Connect…', WorkFlow.Connect.getFunctionName())
    .addItem(`Update…`, WorkFlow.Update.getFunctionName())
    .addItem(`Disconnect…`, WorkFlow.Disconnect.getFunctionName())
    .addToUi();
}

export const onInstall = onOpen;

import * as WorkFlow from './Workflow';

// FIXME update list when new sheet activates
export function onOpen() {
    let menu = SpreadsheetApp.getUi()
        .createAddonMenu()
        .addItem('Connect…', WorkFlow.Connect.getFunctionName())
        .addItem(`Update…`, WorkFlow.Update.getFunctionName())
        .addItem(`Disconnect…`, WorkFlow.Disconnect.getFunctionName())
        .addToUi();
}

export const onInstall = onOpen;

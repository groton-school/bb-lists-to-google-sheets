import * as g from '@battis/gas-lighter';
import * as Metadata from '../Metadata';

export const getFunctionName = () => 'disconnect';
global.disconnect = () => {
    SpreadsheetApp.getUi().showModalDialog(
        g.HtmlService.createTemplateFromFile('templates/disconnect').setHeight(150),
        'Disconnect'
    );
};

global.disconnectRemoveMetadata = () => {
    Metadata.removeList();
    Metadata.removeRange();
    Metadata.removeLastUpdated();
};

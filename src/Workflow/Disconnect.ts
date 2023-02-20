import * as g from '@battis/gas-lighter';
import * as Metadata from '../Metadata';

export const getFunctionName = () => 'disconnect';
global.disconnect = () => {
    const list = Metadata.getList();
    if (list) {
        SpreadsheetApp.getUi().showModalDialog(
            g.HtmlService.createTemplateFromFile('templates/disconnect', {
                list,
            }).setHeight(150),
            'Disconnect'
        );
    } else {
        SpreadsheetApp.getUi().showModalDialog(
            g.HtmlService.createTemplateFromFile('templates/error', {
                message:
                    'No Blackbaud list is connected to this sheet as a data source, so nothing can be disconnected.',
            }).setHeight(100),
            'Not Connected'
        );
    }
};

global.disconnectRemoveMetadata = () => {
    Metadata.removeList();
    Metadata.removeRange();
    Metadata.removeLastUpdated();
};

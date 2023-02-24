import g from '@battis/gas-lighter';
import * as Metadata from '../Metadata';

enum Option {
    Cancel = 'Cancel',
    Disconnect = 'Disconnect',
}

export const getFunctionName = () => 'disconnect';
global.disconnect = () => {
    const list = Metadata.getList();
    if (list) {
        g.SpreadsheetApp.Dialog.showModal({
            message: `You are about to remove the developer metadata that connects this sheet
            to "${list.name}" on Blackbaud. You will no longer be able to
            update the data on this sheet directly from Blackbaud. You will need to
            select the existing data and replace it with a new import from Blackbaud
            if you need to get new data.`,
            title: `Disconnect from Blackbaud`,
            buttons: [Option.Cancel, { name: Option.Disconnect, class: 'red' }],
            functionName: 'disconnectResponse',
            height: 150,
        });
    } else {
        g.SpreadsheetApp.Dialog.showModal({
            message:
                'No Blackbaud list is connected to this sheet as a data source, so nothing can be disconnected.',
            title: 'Not Connected',
        });
    }
};

global.disconnectResponse = (response: Option) => {
    if (response == Option.Disconnect) {
        Metadata.removeList();
        Metadata.removeRange();
        Metadata.removeLastUpdated();
    }
    return null;
};

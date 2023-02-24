import g from '@battis/gas-lighter';
import * as Metadata from '../Metadata';
import * as Import from './ImportData';

export const getFunctionName = () => 'update';
global.update = () => {
    const list = Metadata.getList();
    if (list) {
        const thread = Utilities.getUuid();
        SpreadsheetApp.getUi().showModalDialog(
            g.HtmlService.Element.Progress.getHtmlOutput(thread),
            'Updating'
        );
        Import.importData(list, Import.Target.Update, thread);
    } else {
        g.SpreadsheetApp.Dialog.showModal({
            message:
                'No Blackbaud list is connected to this sheet as a data source, so nothing can be updated. Connect a data source to allow for updating.',
            title: 'Not Connected',
        });
    }
};

import * as g from '@battis/gas-lighter';
import * as Metadata from '../Metadata';
import ImportData, { Target } from './ImportData';

export const getFunctionName = () => 'update';
global.update = () => {
    const list = Metadata.getList();
    if (list) {
        const thread = Utilities.getUuid();
        SpreadsheetApp.getUi().showModalDialog(
            g.HtmlService.Element.Progress.getHtmlOutput(thread),
            'Updating'
        );
        ImportData(list, Target.update, thread);
    } else {
        SpreadsheetApp.getUi().showModalDialog(
            g.HtmlService.createTemplateFromFile('templates/error', {
                message:
                    'No Blackbaud list is connected to this sheet as a data source, so nothing can be updated. Connect a data source to allow for updating.',
            }).setHeight(100),
            'Not Connected'
        );
    }
};

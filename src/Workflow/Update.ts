import * as g from '@battis/gas-lighter';
import * as Metadata from '../Metadata';
import ImportData, { Target } from './ImportData';

export const getFunctionName = () => 'update';
global.update = () => {
    const thread = Utilities.getUuid();
    SpreadsheetApp.getUi().showModalDialog(
        g.HtmlService.Element.Progress.getHtmlOutput(thread),
        'Updating'
    );
    ImportData(Metadata.getList(), Target.update, thread);
};

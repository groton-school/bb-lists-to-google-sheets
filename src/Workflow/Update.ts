import * as g from '@battis/gas-lighter';
import * as Metadata from '../Metadata';
import ImportData, { Target } from './ImportData';

export const getFunctionName = () => 'update';
global.update = () => {
    /*
     * FIXME detect if current sheet is actually updatable
     *   Redirect into Connect workflow if not updateable
     */
    const thread = Utilities.getUuid();
    SpreadsheetApp.getUi().showModalDialog(
        g.HtmlService.Element.Progress.getHtmlOutput(thread),
        'Updating'
    );
    ImportData(Metadata.getList(), Target.update, thread);
};

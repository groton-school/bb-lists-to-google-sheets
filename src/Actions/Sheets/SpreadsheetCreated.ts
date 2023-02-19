import * as g from '@battis/gas-lighter';
import * as State from '../../State';
import OpenSpreadsheet from './OpenSpreadsheet';

export function spreadsheetCreatedCard() {
    return g.CardService.newCard({
        header: State.getSpreadsheet().getName(),
        widgets: [
            `The spreadsheet "${State.getSpreadsheet().getName()}" has been created in ${State.getFolder()
                ? `the folder "${State.getFolder().getName()}"`
                : 'your My Drive'
            } and populated with the data in "${State.getList().name
            }" from Blackbaud.`,
            g.CardService.newTextButton({
                text: 'Open Spreadsheet',
                functionName: OpenSpreadsheet,
            }),
        ],
    });
}

export function spreadsheetCreatedAction() {
    return g.CardService.replaceStack(spreadsheetCreatedCard());
}
global.action_sheets_spreadsheetCreated = spreadsheetCreatedAction;
export default 'action_sheets_spreadsheetCreated';

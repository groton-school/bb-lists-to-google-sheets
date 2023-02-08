import { Terse } from '@battis/gas-lighter';
import * as State from '../../State';
import Home from '../App/Home';

export function sheetAppendedCard() {
    return Terse.CardService.newCard({
        header: State.getSheet().getName(),
        widgets: [
            `The sheet "${State.getSheet().getName()}" has been appended to "${State.getSpreadsheet().getName()}" and populated with the data in "${State.getList().name
            }" from Blackbaud.`,
            Terse.CardService.newTextButton({ text: 'Done', functionName: Home }),
        ],
    });
}

export function sheetAppendedAction() {
    return Terse.CardService.replaceStack(sheetAppendedCard());
}
global.action_sheets_sheetAppended = sheetAppendedAction;
export default 'action_sheets_sheetAppended';

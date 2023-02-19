import * as g from '@battis/gas-lighter';
import * as State from '../../State';
import Home from '../App/Home';

export function sheetAppendedCard() {
    return g.CardService.Card.create({
        header: State.getSheet().getName(),
        widgets: [
            `The sheet "${State.getSheet().getName()}" has been appended to "${State.getSpreadsheet().getName()}" and populated with the data in "${State.getList().name
            }" from Blackbaud.`,
            g.CardService.Widget.newTextButton({ text: 'Done', functionName: Home }),
        ],
    });
}

export function sheetAppendedAction() {
    return g.CardService.Navigation.replaceStack(sheetAppendedCard());
}
global.action_sheets_sheetAppended = sheetAppendedAction;
export default 'action_sheets_sheetAppended';

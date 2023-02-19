import * as g from '@battis/gas-lighter';
import * as Sheets from '../../Sheets';
import * as State from '../../State';
import Home from '../App/Home';

export function updatedCard() {
    return g.CardService.Card.create({
        header: `${State.getSheet().getName()} Updated`,
        widgets: [
            `The sheet "${State.getSheet().getName()}" of "${State.getSpreadsheet().getName()}" has been updated with the current data from "${Sheets.Metadata.getList().name
            }" in Blackbaud.`,
            g.CardService.Widget.newTextButton({ text: 'Done', functionName: Home }),
        ],
    });
}

export function updatedAction() {
    return g.CardService.Navigation.pushCard(updatedCard());
}
global.action_sheets_updated = updatedAction;
export default 'action_sheets_updated';

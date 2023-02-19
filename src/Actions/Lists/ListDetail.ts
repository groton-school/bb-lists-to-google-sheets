import * as g from '@battis/gas-lighter';
import * as State from '../../State';
import ImportData from './ImportData';

export function listDetailCard() {
    var buttonNameBasedOnIntent = 'Create Spreadsheet';
    switch (State.getIntent()) {
        case State.Intent.AppendSheet:
            buttonNameBasedOnIntent = 'Append Sheet';
            break;
        case State.Intent.ReplaceSelection:
            buttonNameBasedOnIntent = `Replace ${State.getSelection().getA1Notation()}`;
            break;
    }

    return g.CardService.Card.create({
        header: State.getList().name,
        widgets: [
            g.CardService.Widget.newDecoratedText({
                topLabel: `${State.getList().type} List`,
                text: State.getList().description,
            }),
            g.CardService.Widget.newDecoratedText({
                topLabel: `Created by ${State.getList().created_by} ${new Date(
                    State.getList().created
                ).toLocaleString()}`,
                bottomLabel: `Last modified ${new Date(
                    State.getList().last_modified
                ).toLocaleString()}`,
            }),
            g.CardService.Widget.newTextButton({
                text: buttonNameBasedOnIntent,
                functionName: ImportData,
            }),
        ],
    });
}

export function listDetailAction(arg) {
    State.update(arg);
    return g.CardService.Navigation.pushCard(listDetailCard());
}
global.action_lists_detail = listDetailAction;
export default 'action_lists_detail';

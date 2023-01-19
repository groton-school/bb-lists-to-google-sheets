import { Terse } from '@battis/google-apps-script-helpers';
import State, { Intent } from '../../State';
import ImportData from './ImportData';

export function listDetailCard() {
    var buttonNameBasedOnIntent = 'Create Spreadsheet';
    switch (State.getIntent()) {
        case Intent.AppendSheet:
            buttonNameBasedOnIntent = 'Append Sheet';
            break;
        case Intent.ReplaceSelection:
            buttonNameBasedOnIntent = 'Replace Selection';
            break;
    }

    return Terse.CardService.newCard({
        header: State.getList().name,
        widgets: [
            Terse.CardService.newDecoratedText({
                topLabel: `${State.getList().type} List`,
                text: State.getList().description,
            }),
            Terse.CardService.newDecoratedText({
                topLabel: `Created by ${State.getList().created_by} ${new Date(
                    State.getList().created
                ).toLocaleString()}`,
                bottomLabel: `Last modified ${new Date(
                    State.getList().last_modified
                ).toLocaleString()}`,
            }),
            Terse.CardService.newTextButton({
                text: buttonNameBasedOnIntent,
                functionName: ImportData,
            }),
        ],
    });
}

export function listDetailAction(arg) {
    State.update(arg);
    return Terse.CardService.pushCard(listDetailCard());
}
global.action_lists_detail = listDetailAction;
export default 'action_lists_detail';

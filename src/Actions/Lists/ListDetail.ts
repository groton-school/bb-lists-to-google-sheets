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

    return CardService.newCardBuilder()
        .setHeader(Terse.CardService.newCardHeader(State.getList().name))
        .addSection(
            CardService.newCardSection()
                .addWidget(
                    Terse.CardService.newDecoratedText(
                        `${State.getList().type} List`,
                        State.getList().description
                    )
                )
                .addWidget(
                    Terse.CardService.newDecoratedText(
                        `Created by ${State.getList().created_by} ${new Date(
                            State.getList().created
                        ).toLocaleString()}`,
                        null,
                        `Last modified ${new Date(
                            State.getList().last_modified
                        ).toLocaleString()}`
                    )
                )
                .addWidget(
                    Terse.CardService.newTextButton(buttonNameBasedOnIntent, ImportData)
                )
        )
        .build();
}

export function listDetailAction(arg) {
    State.update(arg);
    return Terse.CardService.pushCard(listDetailCard());
}
global.action_lists_detail = listDetailAction;
export default 'action_lists_detail';

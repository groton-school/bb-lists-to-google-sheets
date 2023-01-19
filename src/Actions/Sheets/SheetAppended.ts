import { Terse } from '@battis/google-apps-script-helpers';
import State from '../../State';
import Home from '../App/Home';

export function sheetAppendedCard() {
    return CardService.newCardBuilder()
        .setHeader(Terse.CardService.newCardHeader(State.getSheet().getName()))
        .addSection(
            CardService.newCardSection()
                .addWidget(
                    Terse.CardService.newTextParagraph(
                        `The sheet "${State.getSheet().getName()}" has been appended to "${State.getSpreadsheet().getName()}" and populated with the data in "${State.getList().name
                        }" from Blackbaud.`
                    )
                )
                .addWidget(Terse.CardService.newTextButton('Done', Home))
        )
        .build();
}

export function sheetAppendedAction() {
    return Terse.CardService.replaceStack(sheetAppendedCard());
}
global.action_sheets_sheetAppended = sheetAppendedAction;
export default 'action_sheets_sheetAppended';

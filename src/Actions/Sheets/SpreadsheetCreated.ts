import { Terse } from '@battis/google-apps-script-helpers';
import State from '../../State';
import OpenSpreadsheet from './OpenSpreadsheet';

export function spreadsheetCreatedCard() {
    return CardService.newCardBuilder()
        .setHeader(
            Terse.CardService.newCardHeader(State.getSpreadsheet().getName())
        )
        .addSection(
            CardService.newCardSection()
                .addWidget(
                    Terse.CardService.newTextParagraph(
                        `The spreadsheet "${State.getSpreadsheet().getName()}" has been created in ${State.getFolder()
                            ? `the folder "${State.getFolder().getName()}"`
                            : 'your My Drive'
                        } and populated with the data in "${State.getList().name
                        }" from Blackbaud.`
                    )
                )
                .addWidget(
                    Terse.CardService.newTextButton('Open Spreadsheet', OpenSpreadsheet)
                )
        )
        .build();
}

export function spreadsheetCreatedAction() {
    return Terse.CardService.replaceStack(spreadsheetCreatedCard());
}
global.action_sheets_spreadsheetCreated = spreadsheetCreatedAction;
export default 'action_sheets_spreadsheetCreated';

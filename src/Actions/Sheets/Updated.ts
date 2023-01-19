import { Terse } from '@battis/google-apps-script-helpers';
import Sheets from '../../Sheets';
import State from '../../State';
import Home from '../App/Home';

export function updatedCard() {
    return CardService.newCardBuilder()
        .setHeader(
            Terse.CardService.newCardHeader(`${State.getSheet().getName()} Updated`)
        )
        .addSection(
            CardService.newCardSection()
                .addWidget(
                    Terse.CardService.newTextParagraph(
                        `The sheet "${State.getSheet().getName()}" of "${State.getSpreadsheet().getName()}" has been updated with the current data from "${Sheets.metadata.get(Sheets.metadata.LIST).name
                        }" in Blackbaud.`
                    )
                )
                .addWidget(Terse.CardService.newTextButton('Done', Home))
        )
        .build();
}

export function updatedAction() {
    return Terse.CardService.pushCard(updatedCard());
}
global.action_sheets_updated = updatedAction;
export default 'action_sheets_updated';

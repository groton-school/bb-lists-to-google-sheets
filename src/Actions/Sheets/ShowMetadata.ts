import { Terse } from '@battis/google-apps-script-helpers';
import Sheets from '../../Sheets';
import State from '../../State';
import Home from '../App/Home';

export function showMetadataCard() {
    return CardService.newCardBuilder()
        .setHeader(Terse.CardService.newCardHeader(State.getSheet().getName()))
        .addSection(
            CardService.newCardSection()
                .addWidget(
                    Terse.CardService.newDecoratedText(
                        `${State.getSheet().getName()} Deeloper Metadata`,
                        JSON.stringify(
                            {
                                [Sheets.metadata.LIST]: Sheets.metadata.get(
                                    Sheets.metadata.LIST
                                ),
                                [Sheets.metadata.RANGE]: Sheets.metadata.get(
                                    Sheets.metadata.RANGE
                                ),
                                [Sheets.metadata.NAME]: Sheets.metadata.get(
                                    Sheets.metadata.NAME
                                ),
                                [Sheets.metadata.LAST_UPDATED]: Sheets.metadata.get(
                                    Sheets.metadata.LAST_UPDATED
                                ),
                            },
                            null,
                            2
                        )
                    )
                )
                .addWidget(Terse.CardService.newTextButton('Done', Home))
        )
        .build();
}

export function showMetadataAction() {
    return Terse.CardService.pushCard(showMetadataCard());
}
global.action_sheets_showMetadata = showMetadataAction;
export default 'action_sheets_showMetadata';

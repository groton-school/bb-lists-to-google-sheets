import { Terse } from '@battis/google-apps-script-helpers';
import Sheets from '../../Sheets';
import State from '../../State';
import Home from '../App/Home';

export function showMetadataCard() {
    return Terse.CardService.newCard({
        header: State.getSheet().getName(),
        widgets: [
            Terse.CardService.newDecoratedText({
                topLabel: `${State.getSheet().getName()} Deeloper Metadata`,
                text: JSON.stringify(
                    {
                        [Sheets.metadata.LIST]: Sheets.metadata.get(Sheets.metadata.LIST),
                        [Sheets.metadata.RANGE]: Sheets.metadata.get(Sheets.metadata.RANGE),
                        [Sheets.metadata.LAST_UPDATED]: Sheets.metadata.get(
                            Sheets.metadata.LAST_UPDATED
                        ),
                    },
                    null,
                    2
                ),
            }),
            Terse.CardService.newTextButton({ text: 'Done', functionName: Home }),
        ],
    });
}

export function showMetadataAction() {
    return Terse.CardService.pushCard(showMetadataCard());
}
global.action_sheets_showMetadata = showMetadataAction;
export default 'action_sheets_showMetadata';

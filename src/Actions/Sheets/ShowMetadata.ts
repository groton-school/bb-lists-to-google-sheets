import { Terse } from '@battis/gas-lighter';
import * as Sheets from '../../Sheets';
import * as State from '../../State';
import Home from '../App/Home';

export function showMetadataCard() {
    return Terse.CardService.newCard({
        header: State.getSheet().getName(),
        widgets: [
            Terse.CardService.newDecoratedText({
                topLabel: `${State.getSheet().getName()} Deeloper Metadata`,
                text: JSON.stringify(
                    {
                        list: Sheets.Metadata.getList(),
                        range: Sheets.Metadata.getRange(),
                        lastUpdate: Sheets.Metadata.getLastUpdated(),
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

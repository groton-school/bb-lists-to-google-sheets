import * as g from '@battis/gas-lighter';
import * as Sheets from '../../Sheets';
import * as State from '../../State';
import Home from '../App/Home';

export function showMetadataCard() {
    return g.CardService.Card.create({
        header: State.getSheet().getName(),
        widgets: [
            g.CardService.Widget.newDecoratedText({
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
            g.CardService.Widget.newTextButton({ text: 'Done', functionName: Home }),
        ],
    });
}

export function showMetadataAction() {
    return g.CardService.Navigation.pushCard(showMetadataCard());
}
global.action_sheets_showMetadata = showMetadataAction;
export default 'action_sheets_showMetadata';

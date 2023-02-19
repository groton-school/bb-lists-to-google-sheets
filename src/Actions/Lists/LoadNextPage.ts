import * as g from '@battis/gas-lighter';
import * as SKY from '../../SKY';
import * as State from '../../State';
import { insertDataAction } from './InsertData';

export function loadNextPageCard() {
    return g.CardService.newCard({
        header: State.getList().name,
        widgets: [
            'Due to limitations by Blackbaud (a rate-limited API) and by Google (time-limited execution of scripts), human interaction is required to load large data lists from Blackbaud.',
            `${State.getData().length - 1} records have been loaded from "${State.getList().name
            }" so far.`,
            g.CardService.newTextButton({
                text: `Load Page ${State.getPage() + 1}`,
                functionName: 'action_lists_loadNextPage',
            }),
        ],
    });
}

export function loadNextPageAction() {
    State.setPage(State.getPage() + 1);
    const data = (
        SKY.School.Lists.get(
            State.getList().id,
            SKY.ServiceManager.ResponseFormat.Array,
            State.getPage()
        ) as []
    ).slice(1); // trim off unneeded column labels
    State.appendData(data);
    if (data.length == SKY.PAGE_SIZE) {
        return g.CardService.replaceStack(loadNextPageCard());
    } else {
        return insertDataAction();
    }
}
global.action_lists_loadNextPage = loadNextPageAction;
export default 'action_lists_loadNextPage';

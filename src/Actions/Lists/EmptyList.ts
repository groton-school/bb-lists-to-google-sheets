import * as g from '@battis/gas-lighter';
import * as State from '../../State';
import Home from '../App/Home';

export function emptyListCard() {
    return g.CardService.newCard({
        header: State.getList().name,
        widgets: [
            `No data was returned in the list "${State.getList().name
            }" so no sheet was created.`,
            g.CardService.newDecoratedText({
                topLabel: 'List',
                text: JSON.stringify(State.getList(), null, 2),
            }),
            g.CardService.newDecoratedText({
                topLabel: 'Data',
                text: JSON.stringify(State.getData(), null, 2),
            }),
            g.CardService.newTextButton({ text: 'Go Home', functionName: Home }),
        ],
    });
}

export function emptyListAction(data) {
    g.CardService.replaceStack(emptyListCard());
}
global.action_lists_emptyList = emptyListAction;
export default 'action_lists_emptyList';

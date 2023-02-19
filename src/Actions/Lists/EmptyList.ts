import * as g from '@battis/gas-lighter';
import * as State from '../../State';
import Home from '../App/Home';

export function emptyListCard() {
    return g.CardService.Card.create({
        header: State.getList().name,
        widgets: [
            `No data was returned in the list "${State.getList().name
            }" so no sheet was created.`,
            g.CardService.Widget.newDecoratedText({
                topLabel: 'List',
                text: JSON.stringify(State.getList(), null, 2),
            }),
            g.CardService.Widget.newDecoratedText({
                topLabel: 'Data',
                text: JSON.stringify(State.getData(), null, 2),
            }),
            g.CardService.Widget.newTextButton({
                text: 'Go Home',
                functionName: Home,
            }),
        ],
    });
}

export function emptyListAction(data) {
    g.CardService.Navigation.replaceStack(emptyListCard());
}
global.action_lists_emptyList = emptyListAction;
export default 'action_lists_emptyList';

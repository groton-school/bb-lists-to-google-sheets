import * as g from '@battis/gas-lighter';
import * as State from '../../State';
import Home from './Home';

export function stateCard(header = 'Application State') {
    return g.CardService.Card.create({
        header,
        widgets: [
            g.CardService.Widget.newDecoratedText({
                topLabel: 'State',
                text: State.toString(),
            }),
            g.CardService.Widget.newTextButton({ text: 'Home', functionName: Home }),
        ],
    });
}

export function stateAction(header?: string) {
    return g.CardService.Navigation.pushCard(stateCard(header));
}

global.action_app_state = stateAction;
export default 'action_app_state';

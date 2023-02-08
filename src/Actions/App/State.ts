import { Terse } from '@battis/gas-lighter';
import * as State from '../../State';
import Home from './Home';

export function stateCard(header = 'Application State') {
    return Terse.CardService.newCard({
        header,
        widgets: [
            Terse.CardService.newDecoratedText({
                topLabel: 'State',
                text: State.toString(),
            }),
            Terse.CardService.newTextButton({ text: 'Home', functionName: Home }),
        ],
    });
}

export function stateAction(header?: string) {
    return Terse.CardService.pushCard(stateCard(header));
}

global.action_app_state = stateAction;
export default 'action_app_state';

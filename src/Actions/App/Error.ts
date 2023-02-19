import * as g from '@battis/gas-lighter';
import * as State from '../../State';
import Home from './Home';

export function errorCard(header = 'An error occurred', widgets = []) {
    return g.CardService.newCard({
        header,
        sections: [
            g.CardService.newCardSection({
                widgets: [
                    ...widgets,
                    g.CardService.newDecoratedText({
                        topLabel: 'State',
                        text: State.toString(),
                    }),
                    g.CardService.newDecoratedText({
                        topLabel: 'UserProperties',
                        text: JSON.stringify(
                            PropertiesService.getUserProperties().getProperties(),
                            null,
                            2
                        ),
                    }),
                    g.CardService.newTextButton({ text: 'Home', functionName: Home }),
                ],
            }),
        ],
    });
}

export function errorAction(header?: string, widgets = []) {
    return g.CardService.replaceStack(errorCard(header, widgets));
}
global.action_app_error = errorAction;
const Error = 'action_app_error';
export default Error;

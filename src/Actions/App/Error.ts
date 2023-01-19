import { Terse } from '@battis/google-apps-script-helpers';
import State from '../../State';
import Home from './Home';

export function errorCard(header = 'An error occurred', widgets = []) {
    return Terse.CardService.newCard({
        header,
        sections: [
            Terse.CardService.newCardSection({
                widgets: [
                    ...widgets,
                    Terse.CardService.newDecoratedText({
                        topLabel: 'State',
                        text: State.toString(),
                    }),
                    Terse.CardService.newDecoratedText({
                        topLabel: 'UserProperties',
                        text: JSON.stringify(
                            PropertiesService.getUserProperties().getProperties(),
                            null,
                            2
                        ),
                    }),
                    Terse.CardService.newTextButton({ text: 'Home', functionName: Home }),
                ],
            }),
        ],
    });
}

export function errorAction(header?: string, widgets = []) {
    return Terse.CardService.replaceStack(errorCard(header, widgets));
}
global.action_app_error = errorAction;
const Error = 'action_app_error';
export default Error;

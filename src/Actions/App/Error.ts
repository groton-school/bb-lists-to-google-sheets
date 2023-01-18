import { Terse } from '@battis/google-apps-script-helpers';
import State from '../../State';
import Home from './Home';

export function errorCard(message = 'An error occurred') {
    return CardService.newCardBuilder()
        .setHeader(Terse.CardService.newCardHeader(message))
        .addSection(
            CardService.newCardSection()
                .addWidget(
                    Terse.CardService.newDecoratedText(
                        'State',
                        JSON.stringify(
                            {
                                folder: State.getFolder()?.getId(),
                                spreadsheet: State.getSpreadsheet()?.getId(),
                                sheet: State.getSheet()?.getName(),
                                selection: State.getSelection()?.getA1Notation(),
                                list: State.getList(),
                                intent: State.getIntent(),
                                page: State.getPage(),
                                data: State.getData(),
                            },
                            null,
                            2
                        )
                    )
                )
                .addWidget(Terse.CardService.newTextButton('Start Over', Home))
        )
        .build();
}

export function errorAction() {
    return Terse.CardService.replaceStack(errorCard());
}
global.action_app_error = errorAction;
const Error = 'action_app_error';
export default Error;

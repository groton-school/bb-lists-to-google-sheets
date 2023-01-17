import { Terse } from '@battis/google-apps-script-helpers';
import Drive from './Drive';
import Lists from './Lists';
import Sheets from './Sheets';
import State from './State';

export default class App {
    public static launch(event = null) {
        const folder = State.getFolder();
        State.reset();
        if (event) {
            State.setFolder(Drive.inferFolder(event));
        }
        // TODO can also infer folder from parent of current spreadsheet
        return App.cards.home();
    }

    public static actions = class {
        public static home() {
            State.reset();
            return Terse.CardService.replaceStack(App.cards.home());
        }
    };

    public static cards = class {
        public static home() {
            if (State.getSelection()) {
                return Sheets.cards.options();
            }
            return Lists.cards.lists();
        }

        public static error(message = 'An error occurred') {
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
                        .addWidget(
                            Terse.CardService.newTextButton(
                                'Start Over',
                                '__App_actions_home'
                            )
                        )
                )
                .build();
        }
    };
}

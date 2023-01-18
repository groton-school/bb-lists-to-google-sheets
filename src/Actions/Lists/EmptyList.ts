import { Terse } from '@battis/google-apps-script-helpers';
import State from '../../State';
import Home from '../App/Home';

export function emptyListCard() {
    return CardService.newCardBuilder()
        .setHeader(Terse.CardService.newCardHeader(State.getList().name))
        .addSection(
            CardService.newCardSection()
                .addWidget(
                    Terse.CardService.newTextParagraph(
                        JSON.stringify(State.getList(), null, 2)
                    )
                )
                .addWidget(
                    Terse.CardService.newTextParagraph(
                        JSON.stringify(State.getData(), null, 2)
                    )
                )
                .addWidget(
                    Terse.CardService.newTextParagraph(
                        `No data was returned in the list "${State.getList().name
                        }" so no sheet was created.`
                    )
                )
                .addWidget(Terse.CardService.newTextButton('Go Home', Home))
        )
        .build();
}

export function emptyListAction(data) {
    Terse.CardService.replaceStack(emptyListCard());
}
global.action_lists_emptyList = emptyListAction;
const EmptyList = 'action_lists_emptyList';
export default EmptyList;

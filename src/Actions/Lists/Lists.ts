import { Terse } from '@battis/gas-lighter';
import { UNCATEGORIZED } from '../../Constants';
import * as SKY from '../../SKY';
import * as State from '../../State';
import ListDetail from './ListDetail';

export function listsCard() {
    const groupCategories = (categories, list) => {
        if (list.id > 0) {
            if (!list.category) {
                list.category = UNCATEGORIZED;
            }
            if (!categories[list.category]) {
                categories[list.category] = [];
            }
            categories[list.category].push(list);
        }
        return categories;
    };
    const lists = (SKY.School.Lists.get() as SKY.School.Lists.Metadata[]).reduce(
        groupCategories,
        {
            [UNCATEGORIZED]: [],
        }
    );

    var intentBasedActionDescription: string;
    switch (State.getIntent()) {
        case State.Intent.AppendSheet:
            intentBasedActionDescription = `a sheet appended to "${State.getSpreadsheet().getName()}"`;
            break;
        case State.Intent.ReplaceSelection:
            intentBasedActionDescription = `the sheet "${State.getSheet().getName()}", replacing the current selection (${State.getSelection().getA1Notation()})`;
            break;
        case State.Intent.CreateSpreadsheet:
        default:
            intentBasedActionDescription = 'a new spreadsheet';
    }

    const card = CardService.newCardBuilder().addSection(
        Terse.CardService.newCardSection({
            widgets: [
                `Choose the list that you would like to import from Blackbaud into ${intentBasedActionDescription}.`,
            ],
        })
    );

    const sortCategoriesWithUncategorizedLast = (a, b) => {
        if (a == UNCATEGORIZED) {
            return 1;
        } else if (b == UNCATEGORIZED) {
            return -1;
        } else {
            return a.localeCompare(b);
        }
    };

    for (const category of Object.getOwnPropertyNames(lists).sort(
        sortCategoriesWithUncategorizedLast
    )) {
        card.addSection(
            Terse.CardService.newCardSection({
                header: category === UNCATEGORIZED ? 'Uncategorized' : category,
                widgets: lists[category].map((list) =>
                    Terse.CardService.newDecoratedText({
                        text: list.name,
                    }).setOnClickAction(
                        Terse.CardService.newAction({
                            functionName: ListDetail,
                            parameters: {
                                state: { list },
                            },
                        })
                    )
                ),
            })
        );
    }
    return card.build();
}

export function listsAction(arg) {
    State.update(arg);
    return Terse.CardService.pushCard(listsCard());
}
global.action_lists_lists = listsAction;
export default 'action_lists_lists';

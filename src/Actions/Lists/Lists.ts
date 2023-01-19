import { Terse } from '@battis/google-apps-script-helpers';
import Lists from '../../Lists';
import School, { ListMetadata } from '../../SKY/School';
import State, { Intent } from '../../State';
import ListDetail from './ListDetail';

export function listsCard() {
    const groupCategories = (categories, list) => {
        if (list.id > 0) {
            if (!list.category) {
                list.category = Lists.UNCATEGORIZED;
            }
            if (!categories[list.category]) {
                categories[list.category] = [];
            }
            categories[list.category].push(list);
        }
        return categories;
    };
    const lists = (School.lists() as ListMetadata[]).reduce(groupCategories, {
        [Lists.UNCATEGORIZED]: [],
    });

    var intentBasedActionDescription;
    switch (State.getIntent()) {
        case Intent.AppendSheet:
            intentBasedActionDescription = `a sheet appended to "${State.getSpreadsheet().getName()}"`;
            break;
        case Intent.ReplaceSelection:
            intentBasedActionDescription = `the sheet "${State.getSheet().getName()}", replacing the current selection (${State.getSelection().getA1Notation()})`;
            break;
        case Intent.CreateSpreadsheet:
        default:
            intentBasedActionDescription = 'a new spreadsheet';
    }

    const card = CardService.newCardBuilder().addSection(
        CardService.newCardSection().addWidget(
            Terse.CardService.newTextParagraph(
                `Choose the list that you would like to import from Blackbaud into ${intentBasedActionDescription}.`
            )
        )
    );

    const sortCategoriesWithUncategorizedLast = (a, b) => {
        if (a == Lists.UNCATEGORIZED) {
            return 1;
        } else if (b == Lists.UNCATEGORIZED) {
            return -1;
        } else {
            return a.localeCompare(b);
        }
    };

    for (const category of Object.getOwnPropertyNames(lists).sort(
        sortCategoriesWithUncategorizedLast
    )) {
        const section = CardService.newCardSection().setHeader(
            category == Lists.UNCATEGORIZED ? 'Uncategorized' : category
        );
        for (const list of lists[category]) {
            section.addWidget(
                Terse.CardService.newDecoratedText(null, list.name).setOnClickAction(
                    Terse.CardService.newAction(ListDetail, {
                        state: { list },
                    })
                )
            );
        }
        card.addSection(section);
    }
    return card.build();
}

export function listsAction(arg) {
    State.update(arg);
    return Terse.CardService.pushCard(listsCard());
}
global.action_lists_lists = listsAction;
export default 'action_lists_lists';

import { Terse } from '@battis/google-apps-script-helpers';
import { PREFIX } from './Constants';
import Sheets from './Sheets';
import SKY, { Response } from './SKY';
import State, { Intent } from './State';

export default class Lists {
    private static readonly UNCATEGORIZED = `${PREFIX}.Lists.uncategorized`;
    private static BLACKBAUD_PAGE_SIZE = 1000;

    public static setSheetName(sheet, timestamp = null) {
        sheet.setName(
            `${State.getList().name} (${timestamp || new Date().toLocaleString()})`
        );
        Sheets.metadata.set(Sheets.metadata.NAME, sheet.getName(), sheet);
    }

    public static actions = class {
        public static lists() {
            return Terse.CardService.pushCard(Lists.cards.lists());
        }

        public static listDetail() {
            return Terse.CardService.pushCard(Lists.cards.listDetail());
        }

        public static importData() {
            State.setData(SKY.school.v1.lists(State.getList().id, Response.Array));
            if (
                State.getData().length ==
                Lists.BLACKBAUD_PAGE_SIZE + 1 /* column labels */
            ) {
                State.setPage(1);
                return Terse.CardService.replaceStack(Lists.cards.loadNextPage());
            } else {
                return Lists.actions.insertData();
            }
        }

        public static loadNextPage() {
            State.setPage(State.getPage() + 1);
            const data = SKY.school.v1
                .lists(State.getList().id, Response.Array, State.getPage())
                .slice(1); // trim off unneeded column labels
            State.appendData(data);
            if (data.length == Lists.BLACKBAUD_PAGE_SIZE) {
                return Terse.CardService.replaceStack(Lists.cards.loadNextPage());
            } else {
                return Lists.actions.insertData();
            }
        }

        public static insertData() {
            const data = State.getData();

            if (!data || data.length == 0) {
                return Lists.actions.emptyList(
                    SKY.school.v1.lists(State.getList().id, Response.Raw)
                );
            }

            var range = null;
            switch (State.getIntent()) {
                case Intent.AppendSheet:
                    State.setSheet(State.getSpreadsheet().insertSheet());
                    range = Sheets.adjustRange(
                        {
                            row: 1,
                            column: 1,
                            numRows: data.length,
                            numColumns: data[0].length,
                        },
                        null,
                        State.getSheet()
                    );
                    break;
                case Intent.ReplaceSelection:
                    State.getSelection().clearContent();
                    range = Sheets.adjustRange(
                        {
                            row: State.getSelection().getRow(),
                            column: State.getSelection().getColumn(),
                            numRows: data.length,
                            numColumns: data[0].length,
                        },
                        State.getSelection()
                    );
                    break;
                case Intent.UpdateExisting:
                    const metaRange = Sheets.metadata.get(Sheets.metadata.RANGE);
                    range = Sheets.adjustRange(
                        {
                            ...metaRange,
                            numRows: data.length,
                            numColumns: data[0].length,
                        },
                        Sheets.rangeFromJSON(metaRange)
                    );
                    break;
                case Intent.CreateSpreadsheet:
                default:
                    State.setSpreadsheet(
                        SpreadsheetApp.create(
                            State.getList().name,
                            data.length,
                            data[0].length
                        )
                    );
                    if (State.getFolder()) {
                        DriveApp.getFileById(State.getSpreadsheet().getId()).moveTo(
                            State.getFolder()
                        );
                    }
                    State.setSheet(State.getSpreadsheet().getSheets()[0]);
                    range = State.getSheet().getRange(1, 1, data.length, data[0].length);
            }

            range.setValues(data);
            range.offset(0, 0, 1, range.getNumColumns()).setFontWeight('bold');
            const timestamp = new Date().toLocaleString();

            Sheets.metadata.set(
                Sheets.metadata.LIST,
                State.getList(),
                range.getSheet()
            );
            Sheets.metadata.set(
                Sheets.metadata.RANGE,
                Sheets.rangeToJSON(range),
                range.getSheet()
            );
            Sheets.metadata.set(
                Sheets.metadata.LAST_UPDATED,
                timestamp,
                range.getSheet()
            );
            range
                .offset(0, 0, 1, 1)
                .setNote(`Last updated from "${State.getList().name}" ${timestamp}`);

            switch (State.getIntent()) {
                case Intent.ReplaceSelection:
                    Sheets.metadata.set(
                        Sheets.metadata.NAME,
                        `${range.getSheet().getName()}-existing`,
                        range.getSheet()
                    );
                    return Terse.CardService.replaceStack(Sheets.cards.updated());
                case Intent.UpdateExisting:
                    if (
                        range.getSheet().getName() ==
                        Sheets.metadata.get(Sheets.metadata.NAME)
                    ) {
                        Lists.setSheetName(range.getSheet());
                    }
                    return Terse.CardService.replaceStack(Sheets.cards.updated());
                case Intent.AppendSheet:
                    range.getSheet().setFrozenRows(1);
                    Lists.setSheetName(range.getSheet(), timestamp);
                    // TODO why isn't the appended sheet made active?
                    State.getSpreadsheet().setActiveSheet(range.getSheet());
                    return Terse.CardService.replaceStack(Sheets.cards.sheetAppended());
                case Intent.CreateSpreadsheet:
                default:
                    range.getSheet().setFrozenRows(1);
                    Lists.setSheetName(range.getSheet(), timestamp);
                    return Terse.CardService.replaceStack(
                        Sheets.cards.spreadsheetCreated()
                    );
            }
        }

        public static emptyList(data) {
            return Terse.CardService.replaceStack(Lists.cards.emptyList(data));
        }
    };

    public static cards = class {
        public static lists() {
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
            const lists = SKY.school.v1
                .lists()
                .reduce(groupCategories, { [Lists.UNCATEGORIZED]: [] });

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
                        Terse.CardService.newDecoratedText(
                            null,
                            list.name
                        ).setOnClickAction(
                            Terse.CardService.newAction('__Lists_actions_listDetail', {
                                state: { list },
                            })
                        )
                    );
                }
                card.addSection(section);
            }
            return card.build();
        }

        public static listDetail() {
            var buttonNameBasedOnIntent = 'Create Spreadsheet';
            switch (State.getIntent()) {
                case Intent.AppendSheet:
                    buttonNameBasedOnIntent = 'Append Sheet';
                    break;
                case Intent.ReplaceSelection:
                    buttonNameBasedOnIntent = 'Replace Selection';
                    break;
            }

            return CardService.newCardBuilder()
                .setHeader(Terse.CardService.newCardHeader(State.getList().name))
                .addSection(
                    CardService.newCardSection()
                        .addWidget(
                            Terse.CardService.newDecoratedText(
                                `${State.getList().type} List`,
                                State.getList().description
                            )
                        )
                        .addWidget(
                            Terse.CardService.newDecoratedText(
                                `Created by ${State.getList().created_by} ${new Date(
                                    State.getList().created
                                ).toLocaleString()}`,
                                null,
                                `Last modified ${new Date(
                                    State.getList().last_modified
                                ).toLocaleString()}`
                            )
                        )
                        .addWidget(
                            Terse.CardService.newTextButton(
                                buttonNameBasedOnIntent,
                                '__Lists_actions_importData'
                            )
                        )
                )
                .build();
        }

        public static loadNextPage() {
            return CardService.newCardBuilder()
                .setHeader(Terse.CardService.newCardHeader(State.getList().name))
                .addSection(
                    CardService.newCardSection()
                        .addWidget(
                            Terse.CardService.newTextParagraph(
                                'Due to limitations by Blackbaud (a rate-limited API) and by Google (time-limited execution of scripts'
                            )
                        )
                        .addWidget(
                            Terse.CardService.newTextParagraph(
                                `${State.getData().length - 1} records have been loaded from "${State.getList().name
                                }" so far.`
                            )
                        )
                        .addWidget(
                            Terse.CardService.newTextButton(
                                `Load Page ${State.getPage() + 1}`,
                                '__Lists_actions_loadNextPage'
                            )
                        )
                )
                .build();
        }

        public static emptyList(data) {
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
                            Terse.CardService.newTextParagraph(JSON.stringify(data, null, 2))
                        )
                        .addWidget(
                            Terse.CardService.newTextParagraph(
                                `No data was returned in the list "${State.getList().name
                                }" so no sheet was created.`
                            )
                        )
                        .addWidget(
                            Terse.CardService.newTextButton(
                                'Try Another List',
                                '__App_actions_home'
                            )
                        )
                )
                .build();
        }
    };
}

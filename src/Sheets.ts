import { Terse } from '@battis/google-apps-script-helpers';
import App from './App';
import { PREFIX } from './Constants';
import State, { Intent } from './State';

export default class Sheets {
    public static metadata = class {
        public static readonly LIST = `${PREFIX}.list`;
        public static readonly RANGE = `${PREFIX}.range`;
        public static readonly NAME = `${PREFIX}.name`;
        public static readonly LAST_UPDATED = `${PREFIX}.lastUpdated`;

        public static get(key, sheet = null) {
            sheet = sheet || State.getSheet();
            if (sheet) {
                const meta = sheet.createDeveloperMetadataFinder().withKey(key).find();
                if (meta && meta.length) {
                    const value = meta.shift().getValue();
                    try {
                        return JSON.parse(value);
                    } catch (e) {
                        return value;
                    }
                }
            }
            return null;
        }

        public static set(key, value, sheet = null) {
            sheet = sheet || State.getSheet();
            if (sheet) {
                const str = JSON.stringify(value);
                const meta = sheet.createDeveloperMetadataFinder().withKey(key).find();
                if (meta && meta.length) {
                    return meta.shift().setValue(str);
                } else {
                    return sheet.addDeveloperMetadata(key, str);
                }
            }
            return false;
        }

        public static delete(key, sheet = null) {
            sheet = sheet || State.getSheet();
            if (sheet) {
                const meta = sheet.createDeveloperMetadataFinder().withKey(key).find();
                if (meta && meta.length) {
                    return meta.shift().remove();
                }
            }
            return null;
        }
    };

    public static rangeToJSON(range) {
        return {
            row: range.getRow(),
            column: range.getColumn(),
            numRows: range.getNumRows(),
            numColumns: range.getNumColumns(),
            sheet: range.getSheet().getName(),
        };
    }

    public static rangeFromJSON(json) {
        // FIXME this fallback is unsafe without tracking tab changes!
        const sheet =
            State.getSpreadsheet().getSheetByName(json.sheet) || State.getSheet();
        return sheet.getRange(json.row, json.column, json.numRows, json.numColumns);
    }

    public static adjustRange(
        { row, column, numRows, numColumns },
        range = null,
        sheet = null
    ) {
        if (range) {
            sheet = range.getSheet();
            if (numRows > range.getNumRows()) {
                sheet.insertRows(range.getLastRow() + 1, numRows - range.getNumRows());
            }
            if (numColumns > range.getNumColumns()) {
                sheet.insertColumns(
                    range.getLastColumn() + 1,
                    numColumns - range.getNumColumns()
                );
            }
        } else if (sheet) {
            if (numRows < sheet.getMaxRows()) {
                sheet.deleteRows(numRows + 1, sheet.getMaxRows() - numRows);
            }
            if (numColumns < sheet.getMaxColumns()) {
                sheet.deleteColumns(numColumns + 1, sheet.getMaxColumns() - numColumns);
            }
        }
        return sheet.getRange(row, column, numRows, numColumns);
    }

    public static actions = class {
        public static openSpreadsheet() {
            const url = State.getSpreadsheet().getUrl();
            return Terse.CardService.replaceStack(App.launch(), url);
        }

        public static showMetadata() {
            return Terse.CardService.pushCard(Sheets.cards.showMetadata());
        }

        public static breakConnection() {
            return Terse.CardService.pushCard(Sheets.cards.confirmBreakConnection());
        }

        public static deleteMetadata() {
            Sheets.metadata.delete(Sheets.metadata.LIST);
            Sheets.metadata.delete(Sheets.metadata.RANGE);
            Sheets.metadata.delete(Sheets.metadata.NAME);
            Sheets.metadata.delete(Sheets.metadata.LAST_UPDATED);
            return App.actions.home();
        }
    };

    public static cards = class {
        public static options() {
            const card = CardService.newCardBuilder().setHeader(
                Terse.CardService.newCardHeader(
                    `${State.getSpreadsheet().getName()} Options`
                )
            );

            const metaList = Sheets.metadata.get(Sheets.metadata.LIST);
            if (metaList) {
                card.addSection(
                    CardService.newCardSection()
                        .addWidget(
                            Terse.CardService.newDecoratedText(
                                State.getSheet().getName(),
                                `Update the data in the current sheet with the current "${metaList.name}" data from Blackbaud.`,
                                `Last updated ${Sheets.metadata.get(Sheets.metadata.LAST_UPDATED) ||
                                'at an unknown time'
                                }`
                            )
                        )
                        .addWidget(
                            Terse.CardService.newTextParagraph(
                                'If the updated data contains more rows or columns than the current data, rows and/or columns will be added to the right and bottom of the current data to make room for the updated data without overwriting other information on the sheet. If the updated data contains fewer rows or columns than the current data, all non-overwritten rows and/or columns in the current data will be cleared of data.'
                            )
                        )
                        .addWidget(
                            Terse.CardService.newTextButton(
                                'Update',
                                '__Lists_actions_importData',
                                {
                                    state: {
                                        intent: Intent.UpdateExisting,
                                        list: Sheets.metadata.get(Sheets.metadata.LIST),
                                    },
                                }
                            )
                        )
                        .addWidget(
                            Terse.CardService.newTextButton(
                                'Show Metadata',
                                '__Sheets_actions_showMetadata'
                            )
                        )
                        .addWidget(
                            Terse.CardService.newTextButton(
                                'Break Connection',
                                '__Sheets_actions_breakConnection'
                            )
                        )
                );
            } else {
                card.addSection(
                    CardService.newCardSection()
                        .addWidget(
                            Terse.CardService.newDecoratedText(
                                State.getSheet().getName(),
                                `Replace the currently selected cells (${State.getSheet()
                                    .getSelection()
                                    .getActiveRange()
                                    .getA1Notation()}) in the sheet "${State.getSheet().getName()}" with data from Blackbaud`
                            )
                        )
                        .addWidget(
                            Terse.CardService.newTextButton(
                                'Replace Selection',
                                '__Lists_actions_lists',
                                {
                                    state: {
                                        intent: Intent.ReplaceSelection,
                                        selection: State.getSheet().getSelection().getActiveRange(),
                                    },
                                }
                            )
                        )
                );
            }
            return card
                .addSection(
                    CardService.newCardSection()
                        .addWidget(
                            Terse.CardService.newTextButton(
                                'Append New Sheet',
                                '__Lists_actions_lists',
                                {
                                    state: { intent: Intent.AppendSheet },
                                }
                            )
                        )
                        .addWidget(
                            Terse.CardService.newTextButton(
                                'New Spreadsheet',
                                '__Lists_actions_lists',
                                {
                                    state: { intent: Intent.CreateSpreadsheet },
                                }
                            )
                        )
                )
                .build();
        }
        public static sheetAppended() {
            return CardService.newCardBuilder()
                .setHeader(Terse.CardService.newCardHeader(State.getSheet().getName()))
                .addSection(
                    CardService.newCardSection()
                        .addWidget(
                            Terse.CardService.newTextParagraph(
                                `The sheet "${State.getSheet().getName()}" has been appended to "${State.getSpreadsheet().getName()}" and populated with the data in "${State.getList().name
                                }" from Blackbaud.`
                            )
                        )
                        .addWidget(
                            Terse.CardService.newTextButton('Done', '__App_actions_home')
                        )
                )
                .build();
        }
        public static spreadsheetCreated() {
            return CardService.newCardBuilder()
                .setHeader(
                    Terse.CardService.newCardHeader(State.getSpreadsheet().getName())
                )
                .addSection(
                    CardService.newCardSection()
                        .addWidget(
                            Terse.CardService.newTextParagraph(
                                `The spreadsheet "${State.getSpreadsheet().getName()}" has been created in ${State.getFolder()
                                    ? `the folder "${State.getFolder().getName()}"`
                                    : 'your My Drive'
                                } and populated with the data in "${State.getList().name
                                }" from Blackbaud.`
                            )
                        )
                        .addWidget(
                            Terse.CardService.newTextButton(
                                'Open Spreadsheet',
                                '__Sheets_actions_openSpreadsheet'
                            )
                        )
                )
                .build();
        }

        public static updated() {
            return CardService.newCardBuilder()
                .setHeader(
                    Terse.CardService.newCardHeader(
                        `${State.getSheet().getName()} Updated`
                    )
                )
                .addSection(
                    CardService.newCardSection()
                        .addWidget(
                            Terse.CardService.newTextParagraph(
                                `The sheet "${State.getSheet().getName()}" of "${State.getSpreadsheet().getName()}" has been updated with the current data from "${Sheets.metadata.get(Sheets.metadata.LIST).name
                                }" in Blackbaud.`
                            )
                        )
                        .addWidget(
                            Terse.CardService.newTextButton('Done', '__App_actions_home')
                        )
                )
                .build();
        }

        public static showMetadata() {
            return CardService.newCardBuilder()
                .setHeader(Terse.CardService.newCardHeader(State.getSheet().getName()))
                .addSection(
                    CardService.newCardSection()
                        .addWidget(
                            Terse.CardService.newDecoratedText(
                                `${State.getSheet().getName()} Deeloper Metadata`,
                                JSON.stringify(
                                    {
                                        [Sheets.metadata.LIST]: Sheets.metadata.get(
                                            Sheets.metadata.LIST
                                        ),
                                        [Sheets.metadata.RANGE]: Sheets.metadata.get(
                                            Sheets.metadata.RANGE
                                        ),
                                        [Sheets.metadata.NAME]: Sheets.metadata.get(
                                            Sheets.metadata.NAME
                                        ),
                                        [Sheets.metadata.LAST_UPDATED]: Sheets.metadata.get(
                                            Sheets.metadata.LAST_UPDATED
                                        ),
                                    },
                                    null,
                                    2
                                )
                            )
                        )
                        .addWidget(
                            Terse.CardService.newTextButton('Done', '__App_actions_home')
                        )
                )
                .build();
        }

        public static confirmBreakConnection() {
            return CardService.newCardBuilder()
                .setHeader(Terse.CardService.newCardHeader(`Are you sure?`))
                .addSection(
                    CardService.newCardSection()
                        .addWidget(
                            Terse.CardService.newTextParagraph(
                                `You are about to remove the developer metadata that connects this sheet to its Blackbaud data source. You will no longer be able to update the data on this sheet directly from Blackbaud. You will need to select the existing data and replace it with a new import from Blackbaud if you need to get new data.`
                            )
                        )
                        .addWidget(
                            Terse.CardService.newTextButton(
                                'Delete Metadata',
                                '__Sheets_actions_deleteMetadata'
                            )
                        )
                        .addWidget(
                            Terse.CardService.newTextButton('Cancel', '__App_actions_home')
                        )
                )
                .build();
        }
    };
}

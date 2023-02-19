import * as g from '@battis/gas-lighter';
import * as Sheets from '../../Sheets';
import * as State from '../../State';
import ImportData from '../Lists/ImportData';
import Lists from '../Lists/Lists';
import ConfirmBreakConnection from './ConfirmBreakConnection';

export function optionsCard() {
    const card = CardService.newCardBuilder().setHeader(
        g.CardService.newCardHeader(`${State.getSpreadsheet().getName()} Options`)
    );

    const metaList = Sheets.Metadata.getList();
    if (metaList) {
        card.addSection(
            g.CardService.newCardSection({
                widgets: [
                    g.CardService.newDecoratedText({
                        topLabel: State.getSheet().getName(),
                        text: `Update the data in the current sheet with the current "${metaList.name}" data from Blackbaud.`,
                        bottomLabel: `Last updated ${Sheets.Metadata.getLastUpdated() || 'at an unknown time'
                            }`,
                    }),
                    'If the updated data contains more rows or columns than the current data, rows and/or columns will be added to the right and bottom of the current data to make room for the updated data without overwriting other information on the sheet. If the updated data contains fewer rows or columns than the current data, all non-overwritten rows and/or columns in the current data will be cleared of data.',
                    g.CardService.newTextButton({
                        text: 'Update',
                        functionName: ImportData,
                        parameters: {
                            state: {
                                intent: State.Intent.UpdateExisting,
                                list: Sheets.Metadata.getList(),
                            },
                        },
                    }),
                    g.CardService.newTextButton({
                        text: 'Break Connection',
                        functionName: ConfirmBreakConnection,
                    }),
                ],
            })
        );
    } else {
        card.addSection(
            g.CardService.newCardSection({
                widgets: [
                    g.CardService.newDecoratedText({
                        topLabel: State.getSheet().getName(),
                        text: `Select a range in the sheet "${State.getSheet().getName()}" to replace with data from Blackbaud`,
                    }),
                    g.CardService.newTextButton({
                        text: 'Replace Selection',
                        functionName: Lists,
                        parameters: {
                            state: {
                                intent: State.Intent.ReplaceSelection,
                            },
                        },
                    }),
                ],
            })
        );
    }
    return card
        .addSection(
            g.CardService.newCardSection({
                widgets: [
                    g.CardService.newTextButton({
                        text: 'Append New Sheet',
                        functionName: Lists,
                        parameters: {
                            state: { intent: State.Intent.AppendSheet },
                        },
                    }),
                    g.CardService.newTextButton({
                        text: 'New Spreadsheet',
                        functionName: Lists,
                        parameters: {
                            state: { intent: State.Intent.CreateSpreadsheet },
                        },
                    }),
                ],
            })
        )
        .build();
}

export function optionsAction() {
    return g.CardService.replaceStack(optionsCard());
}
global.action_sheets_options = optionsAction;
export default 'action_sheets_options';

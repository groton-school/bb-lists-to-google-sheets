import { Terse } from '@battis/google-apps-script-helpers';
import Sheets from '../../Sheets';
import State, { Intent } from '../../State';
import ImportData from '../Lists/ImportData';
import Lists from '../Lists/Lists';
import ConfirmBreakConnection from './ConfirmBreakConnection';
import ShowMetadata from './ShowMetadata';

export function optionsCard() {
    const card = CardService.newCardBuilder().setHeader(
        Terse.CardService.newCardHeader(
            `${State.getSpreadsheet().getName()} Options`
        )
    );

    const metaList = Sheets.metadata.get(Sheets.metadata.LIST);
    if (metaList) {
        card.addSection(
            Terse.CardService.newCardSection({
                widgets: [
                    Terse.CardService.newDecoratedText({
                        topLabel: State.getSheet().getName(),
                        text: `Update the data in the current sheet with the current "${metaList.name}" data from Blackbaud.`,
                        bottomLabel: `Last updated ${Sheets.metadata.get(Sheets.metadata.LAST_UPDATED) ||
                            'at an unknown time'
                            }`,
                    }),
                    'If the updated data contains more rows or columns than the current data, rows and/or columns will be added to the right and bottom of the current data to make room for the updated data without overwriting other information on the sheet. If the updated data contains fewer rows or columns than the current data, all non-overwritten rows and/or columns in the current data will be cleared of data.',
                    Terse.CardService.newTextButton({
                        text: 'Update',
                        functionName: ImportData,
                        parameters: {
                            state: {
                                intent: Intent.UpdateExisting,
                                list: Sheets.metadata.get(Sheets.metadata.LIST),
                            },
                        },
                    }),
                    Terse.CardService.newTextButton({
                        text: 'Show Metadata',
                        functionName: ShowMetadata,
                    }),
                    Terse.CardService.newTextButton({
                        text: 'Break Connection',
                        functionName: ConfirmBreakConnection,
                    }),
                ],
            })
        );
    } else {
        card.addSection(
            Terse.CardService.newCardSection({
                widgets: [
                    Terse.CardService.newDecoratedText({
                        topLabel: State.getSheet().getName(),
                        text: `Replace the currently selected cells (${State.getSheet()
                            .getSelection()
                            .getActiveRange()
                            .getA1Notation()}) in the sheet "${State.getSheet().getName()}" with data from Blackbaud`,
                    }),
                    Terse.CardService.newTextButton({
                        text: 'Replace Selection',
                        functionName: Lists,
                        parameters: {
                            state: {
                                intent: Intent.ReplaceSelection,
                                selection: State.getSheet().getSelection().getActiveRange(),
                            },
                        },
                    }),
                ],
            })
        );
    }
    return card
        .addSection(
            Terse.CardService.newCardSection({
                widgets: [
                    Terse.CardService.newTextButton({
                        text: 'Append New Sheet',
                        functionName: Lists,
                        parameters: {
                            state: { intent: Intent.AppendSheet },
                        },
                    }),
                    Terse.CardService.newTextButton({
                        text: 'New Spreadsheet',
                        functionName: Lists,
                        parameters: {
                            state: { intent: Intent.CreateSpreadsheet },
                        },
                    }),
                ],
            })
        )
        .build();
}

export function optionsAction() {
    return Terse.CardService.replaceStack(optionsCard());
}
global.action_sheets_options = optionsAction;
export default 'action_sheets_options';

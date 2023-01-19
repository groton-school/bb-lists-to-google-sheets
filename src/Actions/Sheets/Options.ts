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
                    Terse.CardService.newTextButton('Update', ImportData, {
                        state: {
                            intent: Intent.UpdateExisting,
                            list: Sheets.metadata.get(Sheets.metadata.LIST),
                        },
                    })
                )
                .addWidget(
                    Terse.CardService.newTextButton('Show Metadata', ShowMetadata)
                )
                .addWidget(
                    Terse.CardService.newTextButton(
                        'Break Connection',
                        ConfirmBreakConnection
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
                    Terse.CardService.newTextButton('Replace Selection', Lists, {
                        state: {
                            intent: Intent.ReplaceSelection,
                            selection: State.getSheet().getSelection().getActiveRange(),
                        },
                    })
                )
        );
    }
    return card
        .addSection(
            CardService.newCardSection()
                .addWidget(
                    Terse.CardService.newTextButton('Append New Sheet', Lists, {
                        state: { intent: Intent.AppendSheet },
                    })
                )
                .addWidget(
                    Terse.CardService.newTextButton('New Spreadsheet', Lists, {
                        state: { intent: Intent.CreateSpreadsheet },
                    })
                )
        )
        .build();
}

export function optionsAction() {
    return Terse.CardService.replaceStack(optionsCard());
}
global.action_sheets_options = optionsAction;
export default 'action_sheets_options';

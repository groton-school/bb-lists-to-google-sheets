// TODO implement editor add-on to detect when active tab changes
const Sheets = {
  META: {
    LIST: `${App.PREFIX}.list`,
    RANGE: `${App.PREFIX}.range`,
    NAME: `${App.PREFIX}.name`
  },

  rangeToJSON(range) {
    return {
      row: range.getRow(),
      column: range.getColumn(),
      numRows: range.getNumRows(),
      numColumns: range.getNumColumns(),
      sheet: range.getSheet().getName()
    };
  },

  rangeFromJSON(json) {
    return State.spreadsheet
      .getSheetByName(json.sheet)
      .getRange(json.row, json.column, json.numRows, json.numColumns);
  },

  adjustRange({ row, column, numRows, numColumns }, range = null, sheet = null) {
    if (range) {
      sheet = range.getSheet();
      if (numRows > range.getNumRows()) {
        sheet.insertRows(range.getLastRow() + 1, numRows - range.getNumRows());
      }
      if (numColumns > range.getNumColumns()) {
        sheet.insertColumns(range.getLastColumn() + 1, numColumns - range.getNumColumns());
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
  },

  actions: {
    spreadsheetCreated({ parameters: { state } }) {
      State.restore(state);
      const url = State.spreadsheet.getUrl(); // App launch will reset the State
      return TerseCardService.replaceStack(App.launch(), url);
    },

    breakConnection({ parameters: { state } }) {
      State.restore(state);
      return TerseCardService.pushCard(Sheets.cards.confirmBreakConnection());
    },

    deleteMetadata({ parameters: { state } }) {
      State.restore(state);
      for (const meta of State.sheet.getDeveloperMetadata()) {
        switch (meta.getKey()) {
          case Sheets.META.LIST:
          case Sheets.META.NAME:
          case Sheets.META.RANGE:
            meta.remove();
        }
      }
      return App.actions.home();
    }
  },

  cards: {
    options() {
      const card = CardService.newCardBuilder()
        .setHeader(TerseCardService.newCardHeader(`${State.spreadsheet.getName()} Options`));

      if (State.metadata && State.metadata.list) {
        card.addSection(CardService.newCardSection()
          .addWidget(TerseCardService.newDecoratedText(
            State.sheet.getName(),
            `Update the data in the current sheet with the current "${State.metadata.list.name}" data from Blackbaud.`
          ))
          .addWidget(TerseCardService.newTextParagraph(
            'If the updated data contains more rows or columns than the current data, rows and/or columns will be added to the right and bottom of the current data to make room for the updated data without overwriting other information on the sheet. If the updated data contains fewer rows or columns than the current data, all non-overwritten rows and/or columns in the current data will be cleared of data.'
          ))
          .addWidget(TerseCardService.newTextButton('Update', '__Lists_actions_importData', {
            intent: Intent.UpdateExisting,
            list: State.metadata.list
          }))
          .addWidget(TerseCardService.newTextButton(
            'Break Connection',
            '__Sheets_actions_breakConnection'
          )));
      } else {
        card.addSection(CardService.newCardSection()
          .addWidget(TerseCardService.newDecoratedText(
            State.sheet.getName(),
            `Replace the currently selected cells (${State.sheet.getSelection().getActiveRange().getA1Notation()}) in the sheet "${State.sheet.getName()}" with data from Blackbaud`
          ))
          .addWidget(TerseCardService.newTextButton(
            'Replace Selection',
            '__Lists_actions_lists',
            {
              intent: Intent.ReplaceSelection,
              selection: State.sheet.getSelection().getActiveRange()
            }
          ))
        );
      }

      return card
        .addSection(CardService.newCardSection()
          .addWidget(TerseCardService.newTextButton(
            'Append New Sheet',
            '__Lists_actions_lists',
            { intent: Intent.AppendSheet }
          ))
          .addWidget(TerseCardService.newTextButton(
            'New Spreadsheet',
            '__Lists_actions_lists',
            { intent: Intent.CreateSpreadsheet }
          ))
        )
        .build();
    },

    sheetAppended() {
      return CardService.newCardBuilder()
        .setHeader(TerseCardService.newCardHeader(State.sheet.getName()))
        .addSection(CardService.newCardSection()
          .addWidget(TerseCardService.newTextParagraph(
            `The sheet "${State.sheet.getName()}" has been appended to "${State.spreadsheet.getName()}" and populated with the data in "${State.list.name}" from Blackbaud.`
          ))
          .addWidget(TerseCardService.newTextButton('Done', '__App_actions_home')))
        .build();
    },

    spreadsheetCreated() {
      return CardService.newCardBuilder()
        .setHeader(TerseCardService.newCardHeader(State.spreadsheet.getName()))
        .addSection(CardService.newCardSection()
          .addWidget(TerseCardService.newTextParagraph(
            `The spreadsheet "${State.spreadsheet.getName()}" has been created in ${State.folder ? `the folder "${State.folder.getName()}"` : "your My Drive"} and populated with the data in "${State.list.name}" from Blackbaud.`
          ))
          .addWidget(TerseCardService.newTextButton(
            'Open Spreadsheet',
            '__Sheets_actions_spreadsheetCreated'
          )))
        .build();
    },

    updated() {
      if (arguments && arguments.length > 0 && arguments[0].state) {
        State.restore(arguments[0].state);
      }
      return CardService.newCardBuilder()
        .setHeader(TerseCardService.newCardHeader(`${State.sheet.getName()} Updated`))
        .addSection(CardService.newCardSection()
          .addWidget(TerseCardService.newTextParagraph(`The sheet "${State.sheet.getName()}" of "${State.spreadsheet.getName()}" has been updated with the current data from "${State.metadata.list.name}" in Blackbaud.`))
          .addWidget(TerseCardService.newTextButton('Done', '__App_actions_home')))
        .build();
    },

    confirmBreakConnection() {
      return CardService.newCardBuilder()
        .setHeader(TerseCardService.newCardHeader(`Are you sure?`))
        .addSection(CardService.newCardSection()
          .addWidget(TerseCardService.newTextParagraph(`You are about to remove the developer metadata that connects this sheet to its Blackbaud data source. You will no longer be able to update the data on this sheet directly from Blackbaud. You will need to select the existing data and replace it with a new import from Blackbaud if you need to get new data.`))
          .addWidget(TerseCardService.newTextButton('Delete Metadata', '__Sheets_actions_deleteMetadata'))
          .addWidget(TerseCardService.newTextButton('Cancel', '__App_actions_home')))
        .build();
    }
  }
}

function __Sheets_actions_update(...args) {
  return Sheets.actions.update(...args);
}

function __Sheets_actions_spreadsheetCreated(...args) {
  return Sheets.actions.spreadsheetCreated(...args);
}

function __Sheets_actions_breakConnection(...args) {
  return Sheets.actions.breakConnection(...args);
}

function __Sheets_actions_deleteMetadata(...args) {
  return Sheets.actions.deleteMetadata(...args);
}

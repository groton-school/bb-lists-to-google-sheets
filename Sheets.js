const Sheets = {
  META: {
    LIST: 'org.groton.BbListsToGoogleSheets.list',
    RANGE: 'org.groton.BbListsToGoogleSheets.range',
    NAME: 'org.groton.BbListsToGoogleSheets.name'
  },

  actions: {
    replaceSelection({ parameters: { state } }) {
      State.restore(state);
      State.selection = State.sheet.getSelection().getActiveRange();
      State.intent = Intent.ReplaceSelection;

      return CardService.newActionResponseBuilder()
        .setNavigation(CardService.newNavigation()
          .pushCard(Lists.cards.lists()))
        .build();
    },

    appendSheet({ parameters: { state } }) {
      State.restore(state);
      // FIXME insert logic to detect change of sheet on append sheet
      State.intent = Intent.AppendSheet;
      return CardService.newActionResponseBuilder()
        .setNavigation(CardService.newNavigation()
          .pushCard(Lists.cards.lists()))
        .build();
    },

    newSpreadsheet({ parameters: { state } }) {
      State.restore(state);
      // FIXME insert logic to detect change of sheet on new spreadsheet
      State.intent = Intent.CreateSpreadsheet;
      return CardService.newActionResponseBuilder()
        .setNavigation(CardService.newNavigation()
          .pushCard(Lists.cards.lists()))
        .build();
    },

    update({ parameters: { state } }) {
      State.restore(state);
      // FIXME insert logic to detect change of sheet on update
      // FIXME merge with Lists.actions.insertData() logic
      const data = SKY.school.v1.lists(State.metadata.list.id, SKY.Response.Array);
      const updatedRange = [State.metadata.range[0], State.metadata.range[1], data.length, data[0].length];
      if (data.length > State.metadata.range[2]) {
        State.sheet.insertRows(State.metadata.range[0] + State.metadata.range[2], data.length - State.metadata.range[2]);
      }
      if (data[0].length > State.metadata.range[3]) {
        State.sheet.insertColumns(State.metadata.range[1] + State.metadata.range[3], data[0].length - State.metadata.range[3]);
      }
      State.sheet.getRange(...State.metadata.range).clearContent();
      State.sheet.getRange(...updatedRange).setValues(data);

      if (State.sheet.getName() == State.metadata.name) {
        State.sheet.setName(`${State.metadata.list.name} (${new Date().toLocaleString()})`);
        State.sheet.addDeveloperMetadata(Sheets.META.NAME, JSON.stringify(State.sheet.getName()));
      }

      return CardService.newActionResponseBuilder()
        .setNavigation(CardService.newNavigation()
          .pushCard(Sheets.cards.updated()))
        .build();
    },

    spreadsheetCreated({ parameters: { state } }) {
      State.restore(state);
      return App.actions.home({ url: State.spreadsheet.getUrl() })
    },

    breakConnection({ parameters: { state } }) {
      State.restore(state);
      return CardService.newActionResponseBuilder()
        .setNavigation(CardService.newNavigation()
          .pushCard(Sheets.cards.confirmBreakConnection()))
        .build();
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
        .setHeader(CardService.newCardHeader()
          .setTitle(`${State.spreadsheet.getName()} Options`));

      if (State.metadata && State.metadata.list) {
        card.addSection(CardService.newCardSection()
          .addWidget(CardService.newDecoratedText()
            .setTopLabel(State.sheet.getName())
            .setText(`Update the data in the current sheet with the current "${State.metadata.list.name}" data from Blackbaud.`)
            .setWrapText(true))
          .addWidget(CardService.newTextParagraph()
            .setText('If the updated data contains more rows or columns than the current data, rows and/or columns will be added to the right and bottom of the current data to make room for the updated data without overwriting other information on the sheet. If the updated data contains fewer rows or columns than the current data, all non-overwritten rows and/or columns in the current data will be cleared of data.'))
          .addWidget(CardService.newTextButton()
            .setText('Update')
            .setOnClickAction(CardService.newAction()
              .setFunctionName('__Sheets_actions_update')
              .setParameters({ state: State.toJSON() })))
          .addWidget(CardService.newTextButton()
            .setText('Break Connection')
            .setOnClickAction(CardService.newAction()
              .setFunctionName('__Sheets_actions_breakConnection')
              .setParameters({ state: State.toJSON() }))));
      } else {
        card.addSection(CardService.newCardSection()
          .addWidget(CardService.newDecoratedText()
            .setTopLabel(State.sheet.getName())
            .setText(`Replace the currently selected cells (${State.sheet.getSelection().getActiveRange().getA1Notation()}) in the sheet "${State.sheet.getName()}" with data from Blackbaud`)
            .setWrapText(true))
          .addWidget(CardService.newTextButton()
            .setText('Replace Selection')
            .setOnClickAction(CardService.newAction()
              .setMethodName('__Sheets_actions_replaceSelection')
              .setParameters({ state: State.toJSON() }))));
      }

      return card
        .addSection(CardService.newCardSection()
          .addWidget(CardService.newTextButton()
            .setText('Append New Sheet')
            .setOnClickAction(CardService.newAction()
              .setFunctionName('__Sheets_actions_appendSheet')
              .setParameters({ state: State.toJSON() })))
          .addWidget(CardService.newTextButton()
            .setText('New Spreadsheet')
            .setOnClickAction(CardService.newAction()
              .setFunctionName('__Sheets_actions_newSpreadsheet')
              .setParameters({ state: State.toJSON() }))))
        .build();
    },

    spreadsheetCreated() {
      return CardService.newCardBuilder()
        .setHeader(CardService.newCardHeader()
          .setTitle(State.spreadsheet.getName()))
        .addSection(CardService.newCardSection()
          .addWidget(CardService.newTextParagraph()
            .setText(`The spreadsheet "${State.spreadsheet.getName()}"" has been created in the folder "${State.folder.getName()}" and populated with the data in "${State.list.name}" from Blackbaud.`))
          .addWidget(CardService.newTextButton()
            .setText('Open Spreadsheet')
            .setOnClickAction(CardService.newAction()
              .setFunctionName('__Sheets_actions_newSpreadsheet')
              .setParameters({ state: State.toJSON() }))))
        .build();
    },

    updated() {
      if (arguments && arguments.length > 0 && arguments[0].state) {
        State.restore(arguments[0].state);
      }
      return CardService.newCardBuilder()
        .setHeader(CardService.newCardHeader()
          .setTitle(`${State.sheet.getName()} Updated`))
        .addSection(CardService.newCardSection()
          .addWidget(CardService.newTextParagraph()
            .setText(`The sheet "${State.sheet.getName()}" of "${State.spreadsheet.getName()}" has been updated with the current data from "${State.metadata.list.name}" in Blackbaud.`))
          .addWidget(CardService.newTextButton()
            .setText('Done')
            .setOnClickAction(CardService.newAction()
              .setFunctionName('__App_actions_home'))))
        .build();
    },

    confirmBreakConnection() {
      return CardService.newCardBuilder()
        .setHeader(CardService.newCardHeader()
          .setTitle(`Are you sure?`))
        .addSection(CardService.newCardSection()
          .addWidget(CardService.newTextParagraph()
            .setText(`You are about to remove the developer metadata that connects this sheet to its Blackbaud data source. You will no longer be able to update the data on this sheet directly from Blackbaud. You will need to select the existing data and replace it with a new import from Blackbaud if you need to get new data.`))
          .addWidget(CardService.newTextButton()
            .setText('Delete Metadata')
            .setOnClickAction(CardService.newAction()
              .setFunctionName('__Sheets_actions_deleteMetadata')
              .setParameters({ state: State.toJSON() })))
          .addWidget(CardService.newTextButton()
            .setText('Cancel')
            .setOnClickAction(CardService.newAction()
              .setFunctionName('__App_actions_home'))))
        .build();
    }
  }
}

function __Sheets_actions_replaceSelection(...args) {
  return Sheets.actions.replaceSelection(...args);
}

function __Sheets_actions_appendSheet(...args) {
  return Sheets.actions.appendSheet(...args);
}

function __Sheets_actions_newSpreadsheet(...args) {
  return Sheets.actions.newSpreadsheet(...args);
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

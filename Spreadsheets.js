function cardSpreadsheetOption() {
  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle(`${State.spreadsheet.getName()} Options`));

  if (State.metadata && State.metadata.list) {
    card.addSection(CardService.newCardSection()
      .addWidget(CardService.newDecoratedText()
        .setTopLabel(State.sheet.getName())
        .setText(`Update the data in the current sheet with the current "${State.metadata.list.name}" data from Blackbaud.`)
        .setWrapText(true))
      .addWidget(CardService.newTextButton()
        .setText('Update')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('actionUpdateSheet')
          .setParameters({ state: State.toJSON() }))));
  }

  return card
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newTextButton()
        .setText('Append New Sheet')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('actionAppendNewSheet')
          .setParameters({ state: State.toJSON() })))
      .addWidget(CardService.newTextButton()
        .setText('New Spreadsheet')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('actionNewSpreadsheet')
          .setParameters({ state: State.toJSON() }))))
    .setFixedFooter(fixedFooterReportIssue())
    .build();
}

function actionUpdateSheet({ parameters: { state } }) {
  State.restore(state);
  // FIXME insert logic to detect change of sheet on update
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
    State.sheet.addDeveloperMetadata(META_NAME, JSON.stringify(State.sheet.getName()));
  }

  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation()
      .pushCard(CardService.newCardBuilder()
        .setHeader(CardService.newCardHeader()
          .setTitle(`${State.sheet.getName()} Updated`))
        .addSection(CardService.newCardSection()
          .addWidget(CardService.newTextParagraph()
            .setText(`The sheet "${State.sheet.getName()}" of "${State.spreadsheet.getName()}" has been updated with the current data from "${State.metadata.list.name}" in Blackbaud.`))
          .addWidget(CardService.newTextButton()
            .setText('Done')
            .setOnClickAction(CardService.newAction()
              .setFunctionName('actionHome'))))
        .setFixedFooter(fixedFooterReportIssue())
        .build()))
    .build();
}

function actionAppendNewSheet({ parameters: { state } }) {
  State.restore(state);
  // FIXME insert logic to detect change of sheet on append sheet
  State.intent = Intent.AppendSheet;
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation()
      .pushCard(cardLists()))
    .build();
}

function actionNewSpreadsheet({ parameters: { state } }) {
  State.restore(state);
  // FIXME insert logic to detect change of sheet on new spreadsheet
  State.intent = Intent.CreateSpreadsheet;
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation()
      .pushCard(cardLists()))
    .build();
}

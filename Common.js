function cardLists() {
  const lists = SKY.school.v1.lists().reduce(
    (categories, list) => {
      if (list.id > 0) {
        if (list.category) {
          if (!categories[list.category]) {
            categories[list.category] = [];
          }
          categories[list.category].push(list);
        } else {
          categories[SORT_UNCATEGORIZED].push(list);
        }
      }
      return categories;
    },
    { [SORT_UNCATEGORIZED]: [] }
  );

  const card = CardService.newCardBuilder()
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph()
        .setText(`Choose the list that you would like to import from Blackbaud into ${State.intent == Intent.AppendSheet ? `a sheet appended to "${State.spreadsheet.getName()}"` : `a new spreadsheet`}.`)));

  for (const category of Object.getOwnPropertyNames(lists).sort((a, b) => {
    if (a == SORT_UNCATEGORIZED) {
      return 1;
    } else if (b == SORT_UNCATEGORIZED) {
      return -1;
    } else {
      return a - b;
    }
  })) {
    const section = CardService.newCardSection()
      .setHeader(category == SORT_UNCATEGORIZED ? 'Uncategorized' : category);
    for (const list of lists[category]) {
      section.addWidget(CardService.newDecoratedText()
        .setText(list.name)
        .setOnClickAction(CardService.newAction()
            .setFunctionName('actionListDetail')
            .setParameters({ state: State.toJSON({ list })})));
    }
    card.addSection(section);
  }
  return card.build();
}

function actionListDetail({ parameters: { state } }) {
  State.reset(state);
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation()
      .pushCard(cardListDetail()))
    .build();
}

function cardListDetail() {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle(State.list.name))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newDecoratedText()
        .setTopLabel(`${STATE.list.type} List`)
        .setText(State.list.description)
        .setWrapText(true))
      .addWidget(CardService.newDecoratedText()
        .setTopLabel(`Created by ${State.list.created_by} ${new Date(State.list.created).toLocaleString()}`)
        .setText(' ')
        .setBottomLabel(`Last modified ${new Date(State.list.last_modified).toLocaleString()}`))
      .addWidget(CardService.newTextButton()
        .setText(State.intent == Intent.AppendSheet ? 'Append Sheet' : 'Create Spreadsheet')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('actionImportData')
          .setParameters({state: State.toJSON()}))))
    .build();
}

function actionImportData({parameters: {state}}) {
  State.reset(state);
  const data = SKY.school.v1.lists(State.list.id, SKY.Response.Array);
  if (!data || data.length == 0) {
    return cardEmptyList(data);
  }
  if (State.intent == Intent.AppendSheet) {
    State.sheet = State.spreadsheet.insertSheet();
    if (State.sheet.getMaxRows() > data.length) {
      State.sheet.deleteRows(data.length + 1, State.sheet.getMaxRows() - data.length);
    }
    if (State.sheet.getMaxColumns() > data[0].length) {
      State.sheet.deleteColumns(data[0].length + 1, State.sheet.getMaxColumns() - data[0].length);
    }
  } else {
    State.spreadsheet = SpreadsheetApp.create(
      State.list.name,
      data.length,
      data[0].length
    );
    State.sheet = State.spreadsheet.getSheets()[0];
  }
  
  const range = [1, 1, data.length, data[0].length];
  State.sheet.getRange(...range).setValues(data);
  State.sheet.addDeveloperMetadata(META_LIST, JSON.stringify(State.list));
  State.sheet.addDeveloperMetadata(META_RANGE, JSON.stringify(range));
  State.sheet.getRange(1, 1, 1, data[0].length).setFontWeight('bold');
  State.sheet.setFrozenRows(1);

  if (State.intent == Intent.AppendSheet) {
    return actionHome();
  } else {
    State.reset();
    return CardService.newActionResponseBuilder()
      .setOpenLink(CardService.newOpenLink().setUrl(State.spreadsheet.getUrl()))
      .setNavigation(CardService.newNavigation().popToRoot())
      .build();
  }
}

function cardEmptyList(data) {
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation()
      .popToRoot()
      .pushCard(CardService.newCardBuilder()
        .setHeader(CardService.newCardHeader()
          .setTitle(list.name))
        .addSection(CardService.newCardSection()
          .addWidget(CardService.newTextParagraph()
            .setText(JSON.stringify(State.list)))
          .addWidget(CardService.newTextParagraph()
            .setText(JSON.stringify(data)))
          .addWidget(CardService.newTextParagraph()
            .setText(`No data was returned in the list "${State.list.name}" so no sheet was created.`))
          .addWidget(CardService.newTextButton()
            .setText('Try Another List')
            .setOnClickAction(CardService.newAction()
              .setFunctionName('actionHome'))))
        .build()))
    .build();
}

function actionError({parameters: {state, message}}) {
  State.reset(state);
  return CardService.newActionResponseBuilder()
    .setNavigation(new CardService.newNavigation()
      .popToRoot()
      .pushCard(cardError(message)))
    .build();
}
 
function cardError(message = "An error occurred") {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle(message))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newDecoratedText()
        .setText(JSON.stringify({
          folder: State.folder ? {
            id: State.folder.getId(),
            name: State.folder.getName()
          } : State.folder,
          spreadsheet: State.spreadsheet ? {
            id: State.spreadsheet.getId(),
            name: State.spreadsheet.getName()
          } : State.spreadsheet,
          sheet: State.sheet ? {
            sheetId: State.sheet.getSheetId(),
            name: State.sheet.getName()
          } : State.sheet,
          metadata: State.metadata ? {
            list: State.metadata.list,
            range: State.metadata.range
          } : State.metadata,
          list: State.list,
          intent: Intent.serialize(State.intent)
        }, null, 2)))
      .addWidget(CardService.newTextButton()
        .setText('Start Over')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('actionHome'))))
    .build();
}

function actionHome() {
  State.reset();
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().popToRoot())
    .build();
}

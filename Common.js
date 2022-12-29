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
        .setText(`Choose the list that you would like to import from Blackbaud into ${State.intent == Intent.AppendSheet ?
          `a sheet appended to "${State.spreadsheet.getName()}"` :
          `a new spreadsheet`
          }.`)));

  for (const category of Object.getOwnPropertyNames(lists).sort((a, b) => {
    if (a == SORT_UNCATEGORIZED) {
      return 1;
    } else if (b == SORT_UNCATEGORIZED) {
      return -1;
    } else {
      return a.localeCompare(b);
    }
  })) {
    const section = CardService.newCardSection()
      .setHeader(category == SORT_UNCATEGORIZED ? 'Uncategorized' : category);
    for (const list of lists[category]) {
      section.addWidget(CardService.newDecoratedText()
        .setText(list.name)
        .setOnClickAction(CardService.newAction()
          .setFunctionName('actionListDetail')
          .setParameters({ state: State.toJSON({ list }) })));
    }
    card.addSection(section);
  }
  return card.setFixedFooter(fixedFooterReportIssue())
    .build();
}

function actionListDetail({ parameters: { state } }) {
  State.restore(state);
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
        .setTopLabel(`${State.list.type} List`)
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
          .setParameters({ state: State.toJSON() }))))
    .setFixedFooter(fixedFooterReportIssue())
    .build();
}

function fixedFooterReportIssue() {
  return CardService.newFixedFooter()
    .setPrimaryButton(CardService.newTextButton()
      .setText("Report a Problem")
      .setOpenLink(CardService.newOpenLink()
        .setUrl('https://github.com/groton-school/bb-lists-to-google-sheets/issues')));
}

function actionImportData({ parameters: { state } }) {
  State.restore(state);
  const data = SKY.school.v1.lists(State.list.id, SKY.Response.Array);
  const dimensions = {
    rows: data.length,
    columns: data[0].length
  };
  if (!data || dimensions.columns == 0) {
    return actionEmptyList(data);
  }
  if (State.intent == Intent.AppendSheet) {
    State.sheet = State.spreadsheet.insertSheet();
    if (State.sheet.getMaxRows() > dimensions.rows) {
      State.sheet.deleteRows(dimensions.rows + 1, State.sheet.getMaxRows() - dimensions.rows);
    }
    if (State.sheet.getMaxColumns() > dimensions.columns) {
      State.sheet.deleteColumns(dimensions.columns + 1, State.sheet.getMaxColumns() - dimensions.columns);
    }
  } else {
    State.spreadsheet = SpreadsheetApp.create(
      State.list.name,
      dimensions.rows,
      dimensions.columns
    );

    if (State.folder) {
      // TODO Why doesn't the spreadsheet move to the folder?
      DriveApp.getFileById(State.spreadsheet.getId()).moveTo(State.folder);
    }

    State.sheet = State.spreadsheet.getSheets()[0];

  }

  const range = [1, 1, dimensions.rows, dimensions.columns];
  State.sheet.getRange(...range).setValues(data);
  State.sheet.getRange(1, 1, 1, dimensions.columns).setFontWeight('bold');
  State.sheet.setFrozenRows(1);
  State.sheet.setName(`${State.list.name} (${new Date().toLocaleString()})`)

  State.sheet.addDeveloperMetadata(META_LIST, JSON.stringify(State.list));
  State.sheet.addDeveloperMetadata(META_RANGE, JSON.stringify(range));
  State.sheet.addDeveloperMetadata(META_NAME, JSON.stringify(State.sheet.getName()));

  if (State.intent == Intent.AppendSheet) {
    return actionHome();
  } else {
    const url = State.spreadsheet.getUrl();
    State.reset();
    return CardService.newActionResponseBuilder()
      .setOpenLink(CardService.newOpenLink().setUrl(url))
      .setNavigation(CardService.newNavigation().popToRoot())
      .build();
  }
}

function actionEmptyList(data) {
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
        .setFixedFooter(fixedFooterReportIssue())
        .build()))
    .build();
}

function actionError({ parameters: { state, message } }) {
  State.restore(state);
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
        .setText(JSON.stringify(State.toJSON(), null, 2)))
      .addWidget(CardService.newTextButton()
        .setText('Start Over')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('actionHome'))))
    .setFixedFooter(fixedFooterReportIssue())
    .build();
}

function actionHome() {
  State.reset();
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().popToRoot())
    .build();
}

function cardLists() {
  const groupCategories = (categories, list) => {
    if (list.id > 0) {
      if (!list.category) {
        list.category = SORT_UNCATEGORIZED;
      }
      if (!categories[list.category]) {
        categories[list.category] = [];
      }
      categories[list.category].push(list);
    }
    return categories;
  };
  const lists = SKY.school.v1.lists().reduce(groupCategories, { [SORT_UNCATEGORIZED]: [] });

  var intentBasedActionDescription = 'a new spreadsheet';
  switch (State.intent) {
    case Intent.AppendSheet:
      intentBasedActionDescription = `a sheet appended to "${State.spreadsheet.getName()}"`;
      break;
    case Intent.ReplaceSelection:
      intentBasedActionDescription = `the sheet "${State.sheet.getName()}", replacing the current selection (${State.selection.getA1Notation()})`;
      break;
  }

  const card = CardService.newCardBuilder()
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph()
        .setText(`Choose the list that you would like to import from Blackbaud into ${intentBasedActionDescription}.`)));

  const sortCategoriesWithUncategorizedLast = (a, b) => {
    if (a == SORT_UNCATEGORIZED) {
      return 1;
    } else if (b == SORT_UNCATEGORIZED) {
      return -1;
    } else {
      return a.localeCompare(b);
    }
  };

  for (const category of Object.getOwnPropertyNames(lists).sort(sortCategoriesWithUncategorizedLast)) {
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
  var buttonNameBasedOnIntent = 'Create Spreadsheet';
  switch (State.intent) {
    case Intent.AppendSheet:
      buttonNameBasedOnIntent = 'Append Sheet';
      break;
    case Intent.ReplaceSelection:
      buttonNameBasedOnIntent = 'Replace Selection';
      break;
  }

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
        .setText(buttonNameBasedOnIntent)
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
      .setTextButtonStyle(CardService.TextButtonStyle.FILLED)
      .setBackgroundColor('#aaaaaa')
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

  const range = [1, 1, dimensions.rows, dimensions.columns];
  if (State.intent == Intent.AppendSheet) {
    State.sheet = State.spreadsheet.insertSheet();
    if (State.sheet.getMaxRows() > dimensions.rows) {
      State.sheet.deleteRows(dimensions.rows + 1, State.sheet.getMaxRows() - dimensions.rows);
    }
    if (State.sheet.getMaxColumns() > dimensions.columns) {
      State.sheet.deleteColumns(dimensions.columns + 1, State.sheet.getMaxColumns() - dimensions.columns);
    }
  } else if (State.intent == Intent.ReplaceSelection) {
    State.selection.clearContent();
    if (dimensions.rows > State.selection.getNumRows()) {
      State.sheet.insertRows(State.selection.getLastRow() + 1, dimensions.rows - State.selection.getNumRows());
    }
    if (dimensions.columns > State.selection.getNumColumns()) {
      State.sheet.insertColumns(State.selection.getLastColumn() + 1, dimensions.columns - State.selection.getNumColumns())
    }
    State.sheet.setActiveRange(State.selection.offset(0, 0, dimensions.rows, dimensions.columns));
    State.selection = State.sheet.getActiveRange();
    range[0] = State.selection.getRowIndex();
    range[1] = State.selection.getColumnIndex();
  } else {
    State.spreadsheet = SpreadsheetApp.create(
      State.list.name,
      dimensions.rows,
      dimensions.columns
    );
    if (State.folder) {
      DriveApp.getFileById(State.spreadsheet.getId()).moveTo(State.folder);
    }
    State.sheet = State.spreadsheet.getSheets()[0];
  }

  State.sheet.getRange(...range).setValues(data);
  State.sheet.getRange(...range).offset(0, 0, 1, dimensions.columns).setFontWeight('bold');

  State.sheet.addDeveloperMetadata(META_LIST, JSON.stringify(State.list));
  State.sheet.addDeveloperMetadata(META_RANGE, JSON.stringify(range));

  switch (State.intent) {
    case Intent.ReplaceSelection:
      State.sheet.addDeveloperMetadata(META_NAME, JSON.stringify(`${State.sheet.getName()}-existing`));
      return actionHome({ card: cardUpdated({ state: State.toJSON() }) });
    case Intent.AppendSheet:
      State.sheet.setFrozenRows(1);
      State.sheet.setName(`${State.list.name} (${new Date().toLocaleString()})`)
      State.sheet.addDeveloperMetadata(META_NAME, JSON.stringify(State.sheet.getName()));
      return actionHome();
    case Intent.CreateSpreadsheet:
    default:
      State.sheet.setFrozenRows(1);
      State.sheet.setName(`${State.list.name} (${new Date().toLocaleString()})`)
      State.sheet.addDeveloperMetadata(META_NAME, JSON.stringify(State.sheet.getName()));
      return actionHome({ card: cardNewSpreadsheet() });
  }
}

function cardNewSpreadsheet() {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle(State.spreadsheet.getName()))
      .addSection(CardService.newCardSection()
        .addWidget(CardService.newTextParagraph()
          .setText(`The spreadsheet "${State.spreadsheet.getName()}"" has been created in the folder "${State.folder.getName()}" and populated with the data in "${State.list.name}" from Blackbaud.`))
        .addWidget(CardService.newTextButton()
          .setText('Open Spreadsheet')
          .setOnClickAction(CardService.newAction()
            .setFunctionName('actionNewSpreadsheet')
            .setParameters({ state: State.toJSON() }))))
    .build();
}

function actionNewSpreadsheet({ parameters: { state }}) {
  State.restore(state);
  return actionHome({ url: State.spreadsheet.getUrl() })
}

function actionEmptyList(data) {
  return actionHome({ card: cardEmptyData() });
}

function cardEmptyData() {
  return CardService.newCardBuilder()
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
    .build();
}

function actionError({ parameters: { state, message } }) {
  State.restore(state);
  return actionHome({ card: cardError(message) });
}

function cardError(message = "An error occurred") {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle(message))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newDecoratedText()
        .setTopLabel('State')
        .setText(JSON.stringify(JSON.parse(State.toJSON()), null, 2))
        .setWrapText(true))
      .addWidget(CardService.newTextButton()
        .setText('Start Over')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('actionHome'))))
    .setFixedFooter(fixedFooterReportIssue())
    .build();
}


function actionHome() {
  var url = null;
  var card = null;
  if (arguments && arguments.length > 0) {
    url = arguments[0].url;
    card = arguments[0].card;
  }

  var action = CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation()
      .popToRoot()
      .updateCard(card || launch()));

  /* 
   * FIXME Opening link not working
   *   Related to #1 probably
   *   Order of operations doesn't seem to matter (OpenLink before Navigation or vice versa)
   */
  if (url) {
    action = action.setOpenLink(CardService.newOpenLink()
      .setUrl(url));
  }
  return action.build();
}

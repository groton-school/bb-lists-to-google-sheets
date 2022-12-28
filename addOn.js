const UNCATEGORIZED = "org.groton.BbListstoGoogleSheets.Uncategorized";

var currentFolder = null;

function onHomepage(event) {
  if (event && event.sheets && event.sheets.id) {
    return sheetOptions(event.sheets)
  } else {
    return listOfLists(event);
  }
}

function sheetOptions(sheets) {
  const spreadsheet = SpreadsheetApp.openById(sheets.id);
  const card = CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle(`${spreadsheet.getName()} Options`));

  const sheet = spreadsheet.getActiveSheet();
  const metadata = sheet.getDeveloperMetadata();
  var list;
  var range;
  for(const data of metadata) {
    switch(data.getKey()) {
      case 'list':
        list = data.getValue();
        break;
      case 'range':
        range = data.getValue();
        break;
    }
  }
  if (list && range) {
    const _list = JSON.parse(list);
    card.addSection(CardService.newCardSection()
      .addWidget(CardService.newTextParagraph()
        .setText(`Update the data in the current sheet with the current "${_list.name}" data from Blackbaud.`))
      .addWidget(CardService.newTextButton()
        .setText('Update')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('onUpdateSheet')
          .setParameters({list, range, spreadsheet: spreadsheet.getId().toString(), sheet: sheet.getSheetId().toString()}))));
  }
 
  return card.addSection(CardService.newCardSection()
    .addWidget(CardService.newTextButton()
      .setText('New Sheet')
      .setOnClickAction(CardService.newAction()
        .setFunctionName('onNewSheet'))))
    .build();
}

function onUpdateSheet({parameters: {list, range, spreadsheet, sheet}}) {
  list = JSON.parse(list);
  range = JSON.parse(range);
  const data = SKY.school.v1.lists(list.id, SKYResponse.Array);
  const updatedRange = [range[0], range[1], data.length, data[0].length];
  spreadsheet = SpreadsheetApp.openById(spreadsheet);
  sheet = spreadsheet.getSheets().reduce((r, s) => s.getSheetId() == sheet ? s : r);
  if (data.length > range[2]) {
    sheet.insertRows(range[0] + range[2], data.length - range[2]);
  }
  if (data[0].length > range[3]) {
    sheet.insertColumns(range[1] + range[3], data[0].length - range[3]);
  }
  sheet.getRange(...range).clearContent();
  sheet.getRange(...updatedRange).setValues(data);

  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation()
      .pushCard(CardService.newCardBuilder()
        .setHeader(CardService.newCardHeader()
          .setTitle(`${sheet.getName()} Updated`))
        .addSection(CardService.newCardSection()
          .addWidget(CardService.newTextParagraph()
            .setText(`The sheet "${sheet.getName()}" of "${spreadsheet.getName()}" has been updated with the current data from "${list.name}" in Blackbaud.`))
          .addWidget(CardService.newTextButton()
            .setText('Done')
            .setOnClickAction(CardService.newAction()
              .setFunctionName('onDoneUpdateSheet'))))
        .build()))
    .build();
}

function onDoneUpdateSheet() {
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation()
      .popToRoot())
    .build();
}

function onNewSheet() {
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation()
      .pushCard(listOfLists()))
    .build();
}

function listOfLists() {
  const lists = SKY.school.v1.lists().reduce((categories, list) => {
      if (list.id > 0) {
        if (list.category) {
          if (!categories[list.category]) {
            categories[list.category] = [];
          }
          categories[list.category].push(list);
        } else {
          categories[UNCATEGORIZED].push(list);
        }
      }
      return categories;
    },
    {[UNCATEGORIZED]: []}
  );

  const card = CardService.newCardBuilder();

  for (const category of Object.getOwnPropertyNames(lists).sort((a, b) => {
    if (a == UNCATEGORIZED) {
      return 1;
    } else if (b == UNCATEGORIZED) {
      return -1;
    } else {
      return a - b;
    }
  })) {
    const section = CardService.newCardSection()
      .setHeader(category == UNCATEGORIZED ? 'Uncategorized' : category);
    for (const list of lists[category]) {
      section.addWidget(CardService.newDecoratedText()
        .setText(list.name)
        .setOnClickAction(CardService.newAction()
          .setFunctionName('onClickListName')
          .setParameters({list: JSON.stringify(list)})
        )
      );
    }
    card.addSection(section)
  }
  return card.build();
}

function onClickListName({parameters: {list}}) {
  list = JSON.parse(list);
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation()
      .pushCard(listDetail(list)))
    .build();
}

function listDetail(list) {
  return CardService.newCardBuilder()
    .setHeader(CardService.newCardHeader()
      .setTitle(list.name))
    .addSection(CardService.newCardSection()
      .addWidget(CardService.newDecoratedText()
        .setTopLabel(`${list.type} List`)
        .setText(list.description)
        .setWrapText(true))
      .addWidget(CardService.newDecoratedText()
        .setTopLabel(`Created by ${list.created_by} ${new Date(list.created).toLocaleString()}`)
        .setText(' ')
        .setBottomLabel(`Last modified ${new Date(list.last_modified).toLocaleString()}`))
      .addWidget(CardService.newTextButton()
        .setText('Create Sheet')
        .setOnClickAction(CardService.newAction()
          .setFunctionName('onCreateSheet')
          .setParameters({list: JSON.stringify(list)}))))
    .build();
}

function onCreateSheet({parameters: {list}}) {
  list = JSON.parse(list);
  const data = SKY.school.v1.lists(list.id, SKYResponse.Array);
  if (!data || data.length == 0) {
    return emptyList(list, data);
  }
  const spreadsheet = SpreadsheetApp.create(list.name, data.length, data[0].length);
  const sheet = spreadsheet.getSheets()[0];
  const range = [1, 1, data.length, data[0].length];
  sheet.getRange(...range)
    .setValues(data);
  sheet.addDeveloperMetadata('list', JSON.stringify(list), SpreadsheetApp.DeveloperMetadataVisibility.DOCUMENT);
  sheet.addDeveloperMetadata('range', JSON.stringify(range))
  sheet.getRange(1, 1, 1, data[0].length)
    .setFontWeight('bold');
  sheet.setFrozenRows(1);
  return CardService.newActionResponseBuilder()
    .setOpenLink(CardService.newOpenLink()
      .setUrl(spreadsheet.getUrl()))
    .setNavigation(CardService.newNavigation()
      .popToRoot())
    .build();
}

function onDeveloperMetadataFound(sheet) {

}

function emptyList(list, data = 'foozle') {
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation()
      .popToRoot()
      .pushCard(CardService.newCardBuilder()
        .setHeader(CardService.newCardHeader()
          .setTitle(list.name))
        .addSection(CardService.newCardSection()
          .addWidget(CardService.newTextParagraph()
            .setText(JSON.stringify(list)))
          .addWidget(CardService.newTextParagraph()
            .setText(JSON.stringify(data)))
          .addWidget(CardService.newTextParagraph()
            .setText(`No data was returned in the list "${list.name}" so no sheet was created.`))
          .addWidget(CardService.newTextButton()
            .setText('Try Another List')
            .setOnClickAction(CardService.newAction()
              .setFunctionName('onGoHome'))))
        .build()))
    .build()
}

function onGoHome() {
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation()
      .popToRoot())
    .build();
}
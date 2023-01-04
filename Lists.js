const Lists = {
  UNCATEGORIZED: 'org.groton.BbListsToGoogleSheets.Lists.uncategorized',
  BLACKBAUD_PAGE_SIZE: 1000,

  actions: {
    listDetail({ parameters: { state } }) {
      State.restore(state);
      return CardService.newActionResponseBuilder()
        .setNavigation(CardService.newNavigation()
          .pushCard(Lists.cards.listDetail()))
        .build();
    },

    importData({ parameters: { state } }) {
      State.restore(state);
      State.data = SKY.school.v1.lists(State.list.id, SKY.Response.Array);
      if (State.data.length == Lists.BLACKBAUD_PAGE_SIZE + 1 /* column labels */) {
        State.page = 1;
        return CardService.newActionResponseBuilder()
          .setNavigation(CardService.newNavigation()
            .popToRoot()
            .updateCard(Lists.cards.loadNextPage()))
          .build();
      } else {
        return Lists.actions.insertData({ parameters: { state: State.toJSON() } });
      }
    },

    loadNextPage({ parameters: { state } }) {
      State.restore(state);
      State.page++;
      const data = SKY.school.v1.lists(State.list.id, SKY.Response.Array, State.page).slice(1); // trim off unneeded column labels
      State.data.push(...data);
      if (data.length == Lists.BLACKBAUD_PAGE_SIZE) {
        return CardService.newActionResponseBuilder()
          .setNavigation(CardService.newNavigation()
            .popToRoot()
            .updateCard(Lists.cards.loadNextPage()))
          .build();
      } else {
        return Lists.actions.insertData({ parameters: { state: State.toJSON() } });
      }
    },

    insertData({ parameters: { state } }) {
      State.restore(state);
      const data = State.data;
      const dimensions = {
        rows: data.length,
        columns: data[0].length
      };

      if (!data || dimensions.columns == 0) {
        return Lists.actions.emptyList(data);
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

      State.sheet.addDeveloperMetadata(Sheets.META.LIST, JSON.stringify(State.list));
      State.sheet.addDeveloperMetadata(Sheets.META.RANGE, JSON.stringify(range));

      switch (State.intent) {
        case Intent.ReplaceSelection:
          State.sheet.addDeveloperMetadata(Sheets.META.NAME, JSON.stringify(`${State.sheet.getName()}-existing`));
          return App.actions.home({ card: Sheets.cards.updated({ state: State.toJSON() }) });
        case Intent.AppendSheet:
          State.sheet.setFrozenRows(1);
          State.sheet.setName(`${State.list.name} (${new Date().toLocaleString()})`)
          State.sheet.addDeveloperMetadata(Sheets.META.NAME, JSON.stringify(State.sheet.getName()));
          return App.actions.home();
        case Intent.CreateSpreadsheet:
        default:
          State.sheet.setFrozenRows(1);
          State.sheet.setName(`${State.list.name} (${new Date().toLocaleString()})`)
          State.sheet.addDeveloperMetadata(Sheets.META.NAME, JSON.stringify(State.sheet.getName()));
          return App.actions.home({ card: Sheets.cards.spreadsheetCreated() });
      }
    },

    emptyList(data) {
      return App.actions.home({ card: Lists.cards.emptyList() });
    }
  },

  cards: {
    lists() {
      const groupCategories = (categories, list) => {
        if (list.id > 0) {
          if (!list.category) {
            list.category = Lists.UNCATEGORIZED;
          }
          if (!categories[list.category]) {
            categories[list.category] = [];
          }
          categories[list.category].push(list);
        }
        return categories;
      };
      const lists = SKY.school.v1.lists().reduce(groupCategories, { [Lists.UNCATEGORIZED]: [] });

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
        if (a == Lists.UNCATEGORIZED) {
          return 1;
        } else if (b == Lists.UNCATEGORIZED) {
          return -1;
        } else {
          return a.localeCompare(b);
        }
      };

      for (const category of Object.getOwnPropertyNames(lists).sort(sortCategoriesWithUncategorizedLast)) {
        const section = CardService.newCardSection()
          .setHeader(category == Lists.UNCATEGORIZED ? 'Uncategorized' : category);
        for (const list of lists[category]) {
          section.addWidget(CardService.newDecoratedText()
            .setText(list.name)
            .setOnClickAction(CardService.newAction()
              .setFunctionName('__Lists_actions_listDetail')
              .setParameters({ state: State.toJSON({ list }) })));
        }
        card.addSection(section);
      }
      return card.build();
    },

    listDetail() {
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
              .setFunctionName('__Lists_actions_importData')
              .setParameters({ state: State.toJSON() }))))
        .build();
    },

    loadNextPage() {
      return CardService.newCardBuilder()
        .setHeader(CardService.newCardHeader()
          .setTitle(State.list.name))
        .addSection(CardService.newCardSection()
          .addWidget(CardService.newTextParagraph()
            .setText('Due to limitations by Blackbaud (a rate-limited API) and by Google (time-limited execution of scripts'))
          .addWidget(CardService.newTextParagraph()
            .setText(`${State.data.length - 1} records have been loaded from "${State.list.name}" so far.`))
          .addWidget(CardService.newTextButton()
            .setText(`Load Page ${State.page + 1}`)
            .setOnClickAction(CardService.newAction()
              .setFunctionName('__Lists_actions_loadNextPage')
              .setParameters({ state: State.toJSON() }))))
        .build();
    },

    emptyList() {
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
              .setFunctionName('__App_actions_home'))))
        .build();
    }

  }
}

function __Lists_actions_listDetail(...args) {
  return Lists.actions.listDetail(...args);
}

function __Lists_actions_importData(...args) {
  return Lists.actions.importData(...args);
}

function __Lists_actions_loadNextPage(...args) {
  return Lists.actions.loadNextPage(...args);
}

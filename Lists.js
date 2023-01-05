const Lists = {
  UNCATEGORIZED: `${App.PREFIX}.Lists.uncategorized`,
  BLACKBAUD_PAGE_SIZE: 1000,

  setSheetName(sheet) {
    sheet.setName(`${State.list.name} (${new Date().toLocaleString()})`);
    Sheets.metadata.set(Sheets.metadata.NAME, sheet.getName(), sheet);
  },

  actions: {
    lists({ parameters: { state } }) {
      State.restore(state);
      return TerseCardService.pushCard(Lists.cards.lists());
    },

    listDetail({ parameters: { state } }) {
      State.restore(state);
      return TerseCardService.pushCard(Lists.cards.listDetail());
    },

    importData({ parameters: { state } }) {
      State.restore(state);
      State.data = SKY.school.v1.lists(State.list.id, SKY.Response.Array);
      if (
        State.data.length ==
        Lists.BLACKBAUD_PAGE_SIZE + 1 /* column labels */
      ) {
        State.page = 1;
        return TerseCardService.replaceStack(
          Lists.cards.loadNextPage()
        );
      } else {
        return Lists.actions.insertData({
          parameters: { state: State.toJSON() },
        });
      }
    },

    loadNextPage({ parameters: { state } }) {
      State.restore(state);
      State.page++;
      const data = SKY.school.v1
        .lists(State.list.id, SKY.Response.Array, State.page)
        .slice(1); // trim off unneeded column labels
      State.data.push(...data);
      if (data.length == Lists.BLACKBAUD_PAGE_SIZE) {
        return TerseCardService.replaceStack(
          Lists.cards.loadNextPage()
        );
      } else {
        return Lists.actions.insertData({
          parameters: { state: State.toJSON() },
        });
      }
    },

    insertData({ parameters: { state } }) {
      State.restore(state);
      const data = State.data;

      if (!data || data.length == 0) {
        return Lists.actions.emptyList(
          SKY.school.v1.lists(State.list.id, SKY.Response.Raw)
        );
      }

      var range = null;
      switch (State.intent) {
        case Intent.AppendSheet:
          State.sheet = State.spreadsheet.insertSheet();
          range = Sheets.adjustRange(
            {
              row: 1,
              column: 1,
              numRows: data.length,
              numColumns: data[0].length,
            },
            null,
            State.sheet
          );
          break;
        case Intent.ReplaceSelection:
          State.selection.clearContent();
          range = Sheets.adjustRange(
            {
              row: State.selection.getRow(),
              column: State.selection.getColumn(),
              numRows: data.length,
              numColumns: data[0].length,
            },
            State.selection
          );
          break;
        case Intent.UpdateExisting:
          const metaRange = Sheets.metadata.get(
            Sheets.metadata.RANGE
          );
          range = Sheets.adjustRange(
            {
              ...metaRange,
              numRows: data.length,
              numColumns: data[0].length,
            },
            Sheets.rangeFromJSON(metaRange)
          );
          break;
        case Intent.CreateSpreadsheet:
        default:
          State.spreadsheet = SpreadsheetApp.create(
            State.list.name,
            data.length,
            data[0].length
          );
          // FIXME ...and we're back to not moving the spreadsheet to the current folder, again
          if (State.folder) {
            DriveApp.getFileById(State.spreadsheet.getId()).moveTo(
              State.folder
            );
          }
          State.sheet = State.spreadsheet.getSheets()[0];
          range = State.sheet.getRange(
            1,
            1,
            data.length,
            data[0].length
          );
      }

      range.setValues(data);
      range.offset(0, 0, 1, range.getNumColumns()).setFontWeight('bold');

      // TODO add a user-visible comment noting last update timestamp
      Sheets.metadata.set(
        Sheets.metadata.LIST,
        State.list,
        range.getSheet()
      );
      Sheets.metadata.set(
        Sheets.metadata.RANGE,
        Sheets.rangeToJSON(range),
        range.getSheet()
      );

      switch (State.intent) {
        case Intent.ReplaceSelection:
          Sheets.metadata.set(
            Sheets.metadata.NAME,
            `${range.getSheet().getName()}-existing`,
            range.getSheet()
          );
          return TerseCardService.replaceStack(
            Sheets.cards.updated()
          );
        case Intent.UpdateExisting:
          if (
            range.getSheet().getName() ==
            Sheets.metadata.get(Sheets.metadata.NAME)
          ) {
            Lists.setSheetName(range.getSheet());
          }
          return TerseCardService.replaceStack(
            Sheets.cards.updated()
          );
        case Intent.AppendSheet:
          range.getSheet().setFrozenRows(1);
          Lists.setSheetName(range.getSheet());
          // TODO why isn't the appended sheet made active?
          State.spreadsheet.setActiveSheet(range.getSheet());
          return TerseCardService.replaceStack(
            Sheets.cards.sheetAppended()
          );
        case Intent.CreateSpreadsheet:
        default:
          range.getSheet().setFrozenRows(1);
          Lists.setSheetName(range.getSheet());
          return TerseCardService.replaceStack(
            Sheets.cards.spreadsheetCreated()
          );
      }
    },

    emptyList(data) {
      return TerseCardService.replaceStack(Lists.cards.emptyList(data));
    },
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
      const lists = SKY.school.v1
        .lists()
        .reduce(groupCategories, { [Lists.UNCATEGORIZED]: [] });

      var intentBasedActionDescription;
      switch (State.intent) {
        case Intent.AppendSheet:
          intentBasedActionDescription = `a sheet appended to "${State.spreadsheet.getName()}"`;
          break;
        case Intent.ReplaceSelection:
          intentBasedActionDescription = `the sheet "${State.sheet.getName()}", replacing the current selection (${State.selection.getA1Notation()})`;
          break;
        case Intent.CreateSpreadsheet:
        default:
          intentBasedActionDescription = 'a new spreadsheet';
      }

      const card = CardService.newCardBuilder().addSection(
        CardService.newCardSection().addWidget(
          TerseCardService.newTextParagraph(
            `Choose the list that you would like to import from Blackbaud into ${intentBasedActionDescription}.`
          )
        )
      );

      const sortCategoriesWithUncategorizedLast = (a, b) => {
        if (a == Lists.UNCATEGORIZED) {
          return 1;
        } else if (b == Lists.UNCATEGORIZED) {
          return -1;
        } else {
          return a.localeCompare(b);
        }
      };

      for (const category of Object.getOwnPropertyNames(lists).sort(
        sortCategoriesWithUncategorizedLast
      )) {
        const section = CardService.newCardSection().setHeader(
          category == Lists.UNCATEGORIZED ? 'Uncategorized' : category
        );
        for (const list of lists[category]) {
          section.addWidget(
            TerseCardService.newDecoratedText(
              null,
              list.name
            ).setOnClickAction(
              TerseCardService.newAction(
                '__Lists_actions_listDetail',
                { list }
              )
            )
          );
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
        .setHeader(TerseCardService.newCardHeader(State.list.name))
        .addSection(
          CardService.newCardSection()
            .addWidget(
              TerseCardService.newDecoratedText(
                `${State.list.type} List`,
                State.list.description
              )
            )
            .addWidget(
              TerseCardService.newDecoratedText(
                `Created by ${State.list.created_by} ${new Date(
                  State.list.created
                ).toLocaleString()}`,
                null,
                `Last modified ${new Date(
                  State.list.last_modified
                ).toLocaleString()}`
              )
            )
            .addWidget(
              TerseCardService.newTextButton(
                buttonNameBasedOnIntent,
                '__Lists_actions_importData'
              )
            )
        )
        .build();
    },

    loadNextPage() {
      return CardService.newCardBuilder()
        .setHeader(TerseCardService.newCardHeader(State.list.name))
        .addSection(
          CardService.newCardSection()
            .addWidget(
              TerseCardService.newTextParagraph(
                'Due to limitations by Blackbaud (a rate-limited API) and by Google (time-limited execution of scripts'
              )
            )
            .addWidget(
              TerseCardService.newTextParagraph(
                `${State.data.length - 1
                } records have been loaded from "${State.list.name
                }" so far.`
              )
            )
            .addWidget(
              TerseCardService.newTextButton(
                `Load Page ${State.page + 1}`,
                '__Lists_actions_loadNextPage'
              )
            )
        )
        .build();
    },

    emptyList(data) {
      return CardService.newCardBuilder()
        .setHeader(TerseCardService.newCardHeader(State.list.name))
        .addSection(
          CardService.newCardSection()
            .addWidget(
              TerseCardService.newTextParagraph(
                JSON.stringify(State.list, null, 2)
              )
            )
            .addWidget(
              TerseCardService.newTextParagraph(
                JSON.stringify(data, null, 2)
              )
            )
            .addWidget(
              TerseCardService.newTextParagraph(
                `No data was returned in the list "${State.list.name}" so no sheet was created.`
              )
            )
            .addWidget(
              TerseCardService.newTextButton(
                'Try Another List',
                '__App_actions_home'
              )
            )
        )
        .build();
    },
  },
};

function __Lists_actions_lists(...args) {
  return Lists.actions.lists(...args);
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

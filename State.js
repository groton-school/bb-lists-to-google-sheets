const State = {
  folder: null,
  spreadsheet: null,
  sheet: null,
  selection: null,
  metadata: null,
  list: null,
  data: null,
  page: null,
  intent: Intent.CreateSpreadsheet,

  reset: (serializedState = null) => {
    const previousState = serializedState && JSON.parse(serializedState);

    if (previousState) {
      State.spreadsheet = previousState.spreadsheet && SpreadsheetApp.openById(previousState.spreadsheet);
      State.sheet = State.spreadsheet && State.spreadsheet.getSheetByName(previousState.sheet);
      State.metadata = State.metadata;
    } else {
      State.spreadsheet = SpreadsheetApp.getActive();
      State.sheet = State.spreadsheet.getActiveSheet();
      State.metadata = null;
      const metadata = State.sheet && State.sheet.getDeveloperMetadata();
      if (metadata) {
        State.metadata = {};
        for (const meta of metadata) {
          const value = JSON.parse(meta.getValue());
          switch (meta.getKey()) {
            case Sheets.META.LIST:
              State.metadata.list = value;
              break;
            case Sheets.META.RANGE:
              State.metadata.range = value;
              break;
            case Sheets.META.NAME:
              State.metadata.name = value;
              break;
          }
        }
      }
    }

    State.list = previousState && previousState.list;
    State.intent = previousState && Intent.deserialize(previousState.intent);

    State.data = previousState && previousState.data;
    State.page = previousState && previousState.page;

    State.selection = null;
    if (previousState && previousState.selection) {
      State.selection = SpreadsheetApp.setActiveRange(State.sheet.getRange(previousState.selection));
    }

    State.folder = null;
    if (previousState && previousState.folder) {
      State.folder = DriveApp.getFolderById(previousState.folder);
    }

  },

  restore: (serializedState) => {
    return State.reset(serializedState);
  },

  toJSON: (stateChanges) => {
    return JSON.stringify({
      spreadsheet: State.spreadsheet ? State.spreadsheet.getId() : State.spreadsheet,
      sheet: State.sheet ? State.sheet.getName() : State.sheet,
      metadata: State.metadata,
      folder: State.folder ? State.folder.getId() : State.folder,
      list: State.list,
      data: State.data,
      page: State.page,
      selection: State.selection ? State.selection.getA1Notation() : State.selection,
      intent: Intent.serialize(State.intent),
      ...stateChanges
    });
  }
}

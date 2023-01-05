const State = {
  folder: null,
  spreadsheet: null,
  sheet: null,
  selection: null,
  list: null,
  data: null,
  page: null,
  intent: Intent.CreateSpreadsheet,

  reset: (serializedState = null) => {
    // TODO reduce unnecessary rehydration -- only do so on demand, as-needed
    const previousState = serializedState && JSON.parse(serializedState);

    if (previousState && previousState.spreadsheet) {
      State.spreadsheet =
        previousState.spreadsheet &&
        SpreadsheetApp.openById(previousState.spreadsheet);
      State.sheet =
        State.spreadsheet &&
        State.spreadsheet.getSheetByName(previousState.sheet);
    } else {
      State.spreadsheet = SpreadsheetApp.getActive();
      State.sheet =
        State.spreadsheet && State.spreadsheet.getActiveSheet();
    }

    State.list = previousState && previousState.list;
    State.intent =
      previousState && Intent.deserialize(previousState.intent);
    State.data = previousState && previousState.data;
    State.page = previousState && previousState.page;

    State.selection = null;
    if (previousState && previousState.selection) {
      State.selection = State.sheet.getRange(previousState.selection);
    }

    State.folder = null;
    if (previousState && previousState.folder) {
      State.folder = DriveApp.getFolderById(previousState.folder);
    }
  },

  restore(serializedState) {
    return State.reset(serializedState);
  },

  restore: serializedState => {
    return State.reset(serializedState);
  },

  toJSON: stateChanges => {
    if (stateChanges) {
      for (const key of Object.keys(stateChanges)) {
        State[key] = stateChanges[key];
      }
    }
    return JSON.stringify({
      spreadsheet: State.spreadsheet
        ? State.spreadsheet.getId()
        : State.spreadsheet,
      sheet: State.sheet ? State.sheet.getName() : State.sheet,
      folder: State.folder ? State.folder.getId() : State.folder,
      list: State.list,
      data: State.data,
      page: State.page,
      selection: State.selection
        ? State.selection.getA1Notation()
        : State.selection,
      intent: Intent.serialize(State.intent),
    });
  },
};

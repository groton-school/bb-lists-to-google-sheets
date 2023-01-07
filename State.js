const State = {
  folder: null,
  spreadsheet: null,
  sheet: null,
  selection: null,
  list: null,
  data: null,
  page: null,
  intent: Intent.CreateSpreadsheet,

  _folder: null,
  _spreadsheet: null,
  _sheet: null,
  _selection: null,
  _data: null,
  _intent: null,

  getFolder() {
    if (State.folder) {
      if (!State._folder) {
        State._folder = DriveApp.getFolderById(State.folder);
      }
      return State._folder;
    }
  },

  setFolder(folder) {
    State._folder = folder;
    State.folder = (folder && folder.getId()) || null;
  },

  getSpreadsheet() {
    if (State.spreadsheet) {
      if (!State._spreadsheet) {
        State._spreadsheet = SpreadsheetApp.openById(State.spreadsheet);
      }
      return State._spreadsheet;
    }
  },

  setSpreadsheet(spreadsheet) {
    State._spreadsheet = spreadsheet;
    State.spreadsheet = spreadsheet.getId();
  },

  getSheet() {
    if (State.sheet && State.spreadsheet) {
      if (!State._sheet) {
        State._sheet = State.getSpreadsheet().getSheetByName(
          State.sheet
        );
      }
      return State._sheet;
    }
  },

  setSheet(sheet) {
    State._sheet = sheet;
    State.sheet = sheet.getName();
    State.setSpreadsheet(sheet.getParent());
  },

  getSelection() {
    if (State.selection) {
      if (!State._selection) {
        State._selection = State.getSheet().getRange(State.selection);
      }
      return State._selection;
    }
  },

  setSelection(range) {
    State._selection = range;
    State.selection = range.getA1Notation();
    State.setSheet(range.getSheet());
  },

  getList() {
    return State.list;
  },

  setList(list) {
    State.list = list;
  },

  getData() {
    if (State.data) {
      return State._data;
    }
  },

  _parseData() {
    if (!State._data) {
      State._parseData();
      State._data = JSON.parse(State.data);
    }
  },

  setData(data) {
    State._data = data;
    State.data = true;
  },

  appendData(data) {
    if (State.data) {
      State._parseData();
      State._data.push(...data);
    } else {
      State.setData(data);
    }
  },

  getPage() {
    return State.page;
  },

  setPage(page) {
    State.page = page;
  },

  getIntent() {
    if (State.intent) {
      if (!State._intent) {
        State._intent = Intent.deserialize(State.intent);
      }
      return State._intent;
    }
  },

  setIntent(intent) {
    State._intent = intent;
    State.intent = Intent.serialize(intent);
  },

  reset: (serializedState = null) => {
    const previousState = serializedState && JSON.parse(serializedState);
    if (previousState) {
      for (const key of Object.keys(previousState)) {
        State[key] = previousState[key];
      }
    }

    if (!State.spreadsheet) {
      State.setSelection(
        SpreadsheetApp.getActive()
          .getActiveSheet()
          .getSelection()
          .getActiveRange()
      );
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
        const value = stateChanges[key];
        switch (key) {
          case 'folder':
            State.setFolder(value);
            break;
          case 'spreadsheet':
            State.setSpreadsheet(value);
            break;
          case 'sheet':
            State.setSheet(value);
            break;
          case 'selection':
            State.setSelection(value);
            break;
          case 'intent':
            State.setIntent(value);
            break;
          default:
            State[key] = stateChanges[key];
        }
      }
    }

    return JSON.stringify({
      folder: State.folder,
      spreadsheet: State.spreadsheet,
      sheet: State.sheet,
      selection: State.selection,
      intent: State.intent,
      list: State.list,
      data: State.getData(),
      page: State.page,
    });
  },
};

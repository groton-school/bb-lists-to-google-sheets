const State = {
  folder: null,
  spreadsheet: null,
  sheet: null,
  metadata: null,
  list: null,
  intent: Intent.CreateSpreadsheet,

  reset: (json = null) => {
    const data = json && JSON.parse(json);

    State.spreadsheet = SpreadsheetApp.getActiveSpreadsheet() || null;
    State.sheet = SpreadsheetApp.getActiveSheet() || null;
    State.list = data && data.list;
    State.intent = data && Intent.deserialize(data.intent);

    State.folder = null;
    if (data && data.folder) {
      State.folder = DriveApp.getFolderById(data.folder);
    }

    State.metadata = null;
    const metadata = State.sheet && State.sheet.getDeveloperMetadata();
    if (metadata) {
      try {
        State.metadata = {
          list: JSON.parse(metadata.filter(data => data.getKey() == META_LIST).shift().getValue()),
          range: JSON.parse(metadata.filter(data => data.getKey() == META_RANGE).shift().getValue())
        }
      } catch (e) {}
    }
  },

  inferFolderFromLaunchEvent: (event) => {
    if (event.drive && event.drive && event.drive.activeCursorItem) {
      const file = DriveApp.getFileById(event.drive.activeCursorItem.id);
      const parents = file.getParents();
      return parents.next() || null;
    }
    return null;
  },

  toJSON: (updates) => {
    return JSON.stringify({
      folder: State.folder ? State.folder.getId() : State.folder,
      list: State.list,
      intent: Intent.serialize(State.intent),
      ...updates
    });
  }
}
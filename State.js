const State = {
  folder: null,
  spreadsheet: null,
  sheet: null,
  metadata: null,
  list: null,
  intent: Intent.CreateSpreadsheet,

  reset: (serializedState = null) => {
    const previousState = serializedState && JSON.parse(serializedState);

    State.spreadsheet = SpreadsheetApp.getActiveSpreadsheet() || null;
    State.sheet = SpreadsheetApp.getActiveSheet() || null;
    State.list = previousState && previousState.list;
    State.intent = previousState && Intent.deserialize(previousState.intent);

    State.folder = null;
    if (previousState && previousState.folder) {
      State.folder = DriveApp.getFolderById(previousState.folder);
    }

    State.metadata = null;
    const metadata = State.sheet && State.sheet.getDeveloperMetadata();
    if (metadata) {
      State.metadata = {};
      for (const meta of metadata) {
        const value = JSON.parse(meta.getValue());
        switch (meta.getKey()) {
          case META_LIST:
            State.metadata.list = value;
            break;
          case META_RANGE:
            State.metadata.range = value;
            break;
          case META_NAME:
            State.metadata.name = value;
            break;
        }
      }
    }
  },

  restore: (serializedState) => {
    return State.reset(serializedState);
  },

  inferFolderFromLaunchEvent: (event) => {
    State.folder = null;
    if (event.drive && event.drive && event.drive.activeCursorItem) {
      const file = DriveApp.getFileById(event.drive.activeCursorItem.id);
      const parents = file.getParents();
      State.folder = parents.next() || null;
    }
  },

  toJSON: (stateChanges) => {
    return JSON.stringify({
      /** for debugging only -- will _not_ be deserialized */
      spreadsheet: State.spreadsheet ? State.spreadsheet.getId() : State.spreadsheet,
      sheet: State.sheet ? State.sheet.getSheetId() : State.sheet,
      metadata: State.metadata,

      /* actually deserialized later */
      folder: State.folder ? State.folder.getId() : State.folder,
      list: State.list,
      intent: Intent.serialize(State.intent),
      ...stateChanges
    });
  }
}

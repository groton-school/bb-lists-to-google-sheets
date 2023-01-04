const Intent = {
  CreateSpreadsheet: Symbol('create'),
  AppendSheet: Symbol('append'),
  ReplaceSelection: Symbol('replace'),
  UpdateExisting: Symbol('update'),

  serialize: (sym) => {
    switch (sym) {
      case Intent.CreateSpreadsheet:
        return 'create';
      case Intent.AppendSheet:
        return 'append';
      case Intent.ReplaceSelection:
        return 'replace';
      case Intent.UpdateExisting:
        return 'update';
      default:
        return null;
    }
  },
  deserialize: (str) => {
    switch (str) {
      case 'create':
        return Intent.CreateSpreadsheet;
      case 'append':
        return Intent.AppendSheet;
      case 'replace':
        return Intent.ReplaceSelection;
      case 'update':
        return Intent.UpdateExisting;
      default:
        return null;
    }
  }
}

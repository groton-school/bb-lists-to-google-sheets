const Intent = {
  CreateSpreadsheet: Symbol('create'),
  AppendSheet: Symbol('append'),
  serialize: (sym) => {
    switch(sym) {
      case Intent.CreateSpreadsheet:
        return 'create';
      case Intent.AppendSheet:
        return 'append';
      default:
        return null;
    }
  },
  deserialize: (str) => {
    switch(str) {
      case 'create':
        return Intent.CreateSpreadsheet;
      case 'append':
        return Intent.AppendSheet;
      default:
        return null;
    }
  }
}

import g from '@battis/gas-lighter';
import * as Metadata from '../Metadata';
import * as SKY from '../SKY';

const UNCATEGORIZED = 'Uncategorized';

function groupCategories(
  categories: { [category: string]: SKY.School.Lists.Metadata[] },
  list: SKY.School.Lists.Metadata
) {
  if (list.id > 0) {
    if (!list.category) {
      list.category = UNCATEGORIZED;
    }
    if (!categories[list.category]) {
      categories[list.category] = [];
    }
    categories[list.category].push(list);
  }
  return categories;
}

export const getFunctionName = () => 'connect';
global.connect = () => {
  const prevList = Metadata.getList();
  const selection = SpreadsheetApp.getActive()
    .getActiveSheet()
    .getActiveRange()
    .getA1Notation();
  SpreadsheetApp.getUi().showModalDialog(
    g.HtmlService.Template.createTemplateFromFile('templates/connect', {
      job: Utilities.getUuid(),
      prevList,
      selection
    }),
    'Connect'
  );
};

global.connectGetLists = () =>
  (SKY.School.Lists.get() as SKY.School.Lists.Metadata[]).reduce(
    groupCategories,
    {
      [UNCATEGORIZED]: []
    }
  );

global.getImportTargetOptions = (list: SKY.School.Lists.Metadata) => {
  const prevList = Metadata.getList();
  const selection = SpreadsheetApp.getActive()
    .getActiveSheet()
    .getActiveRange()
    .getA1Notation();
  let message = `Where would you like the data from "${list.name}" to be imported?`;
  const buttons = [
    { name: 'Add Sheet', value: 'sheet', class: 'action' },
    { name: 'New Spreadsheet', value: 'spreadsheet', class: 'action' }
  ];
  if (prevList) {
    message += ` This sheet is already connected to "${prevList.name}" on Blackbaud, so you need to choose a new sheet or spreadsheet as a new destination for "${list.name}"`;
  } else {
    message += ` If you choose to replace ${selection}, its contents will be erased and, if necessary, additional rows and/or columns will be added to make room for the data from "${list.name}".`;
    buttons.unshift({
      name: `Replace ${selection}`,
      value: 'selection',
      class: 'create'
    });
  }
  return g.SpreadsheetApp.Dialog.getHtml({
    message,
    buttons,
    handler: 'handleTargetSubmit'
  });
};

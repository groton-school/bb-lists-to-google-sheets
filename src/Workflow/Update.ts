import * as Metadata from '../Metadata';
import * as Import from './ImportData';
import g from '@battis/gas-lighter';

export const getFunctionName = () => 'update';
global.update = () => {
  const list = Metadata.getList();
  if (list) {
    const progress = new g.HtmlService.Element.Progress();
    progress.showModalDialog(SpreadsheetApp, 'Updating');
    Import.importData(list, Import.Target.Update, progress.job);
  } else {
    g.SpreadsheetApp.Dialog.showModal({
      message:
        'No Blackbaud list is connected to this sheet as a data source, so nothing can be updated. Connect a data source to allow for updating.',
      title: 'Not Connected'
    });
  }
};

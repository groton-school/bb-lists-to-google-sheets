import * as g from '@battis/gas-lighter';
import * as AddOn from './AddOn';

global.onOpen = AddOn.onOpen;
global.onInstall = AddOn.onInstall;

global.include = g.HtmlService.include;
global.getProgress = (thread: string) =>
    g.HtmlService.Element.Progress.getProgress(thread);
global.dialogClose = g.UI.Dialog.close;

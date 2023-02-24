import g from '@battis/gas-lighter';
import * as AddOn from './AddOn';

global.onOpen = AddOn.onOpen;
global.onInstall = AddOn.onInstall;

// TODO g.Globals.register();
global.include = g.HtmlService.include;
global.getProgress = g.HtmlService.Element.Progress.getProgress;
global.dialogClose = g.UI.Dialog.dialogClose;

import * as AddOn from './AddOn';
import g from '@battis/gas-lighter';

global.onOpen = AddOn.onOpen;
global.onInstall = AddOn.onInstall;

global.include = g.HtmlService.Template.include;
global.getProgress = g.HtmlService.Element.Progress.getProgress;
global.dialogClose = g.UI.Dialog.dialogClose;

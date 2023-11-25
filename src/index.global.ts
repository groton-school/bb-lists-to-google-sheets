import * as AddOn from './AddOn';
import { getProgress } from '@battis/gas-lighter/src/HtmlService/Element/Progress/index.global';
import { include } from '@battis/gas-lighter/src/HtmlService/Template/index.global';
import { dialogClose } from '@battis/gas-lighter/src/UI/Dialog';

global.onOpen = AddOn.onOpen;
global.onInstall = AddOn.onInstall;

global.include = include;
global.getProgress = getProgress;
global.dialogClose = dialogClose;

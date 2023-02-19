import * as g from '@battis/gas-lighter';
import * as State from '../../State';
import { homeCard } from '../App/Home';

export function openSpreadsheetAction() {
    const url = State.getSpreadsheet().getUrl();
    return g.CardService.Navigation.replaceStack(homeCard(), url);
}
global.action_sheets_openSpreadsheet = openSpreadsheetAction;
export default 'action_sheets_openSpreadsheet';

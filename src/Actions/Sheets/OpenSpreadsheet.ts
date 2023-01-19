import { Terse } from '@battis/google-apps-script-helpers';
import State from '../../State';
import { homeCard } from '../App/Home';

export function openSpreadsheetAction() {
    const url = State.getSpreadsheet().getUrl();
    return Terse.CardService.replaceStack(homeCard(), url);
}
global.action_sheets_openSpreadsheet = openSpreadsheetAction;
export default 'action_sheets_openSpreadsheet';

import { Terse } from '@battis/gas-lighter';
import * as Drive from '../../Drive';
import * as State from '../../State';
import { listsCard } from '../Lists/Lists';
import { optionsCard } from '../Sheets/Options';

export function homeCard(event?: GoogleAppsScript.Events.AppsScriptEvent) {
    const folder = State.getFolder();
    State.reset();
    if (event) {
        State.setFolder(Drive.inferFolder(event));
    } else if (folder) {
        State.setFolder(folder);
    }
    // TODO can also infer folder from parent of current spreadsheet
    if (State.getSelection()) {
        return optionsCard();
    }
    return listsCard();
}

export function homeAction(event?: GoogleAppsScript.Events.AppsScriptEvent) {
    return Terse.CardService.replaceStack(homeCard());
}

global.action_app_home = homeAction;
export default 'action_app_home';

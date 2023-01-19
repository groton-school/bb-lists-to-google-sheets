import { Terse } from '@battis/google-apps-script-helpers';
import Drive from '../../Drive';
import State from '../../State';
import { listsCard } from '../Lists/Lists';
import { optionsCard } from '../Sheets/Options';
import { errorCard } from './Error';

export function homeCard(event?) {
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
    return errorCard('Debugging', [
        Terse.CardService.newDecoratedText({
            topLabel: 'ActiveRange',
            text: JSON.stringify(
                SpreadsheetApp.getActive().getActiveRangeList(),
                null,
                2
            ),
        }),
    ]);
    return listsCard();
}

export function homeAction(event?) {
    return Terse.CardService.replaceStack(homeCard());
}
global.action_app_home = homeAction;
const Home = 'action_app_home';
export default Home;

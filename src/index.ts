import { homeCard } from './Actions/App/Home';
import Drive from './Drive';
import State from './State';

global.launch = (event) => {
    const folder = State.getFolder();
    State.reset();
    if (event) {
        State.setFolder(Drive.inferFolder(event));
    }
    // TODO can also infer folder from parent of current spreadsheet
    return homeCard();
};

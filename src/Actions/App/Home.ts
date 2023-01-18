import { Terse } from '@battis/google-apps-script-helpers';
import State from '../../State';
import { listsCard } from '../Lists/Lists';
import { optionsCard } from '../Sheets/Options';

export function homeCard() {
    if (State.getSelection()) {
        return optionsCard();
    }
    return listsCard();
}

export function homeAction() {
    State.reset();
    return Terse.CardService.replaceStack(homeCard());
}
global.action_app_home = homeAction;
const Home = 'action_app_home';
export default Home;

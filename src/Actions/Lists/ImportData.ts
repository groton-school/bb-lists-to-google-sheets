import { Terse } from '@battis/gas-lighter';
import * as SKY from '../../SKY';
import * as State from '../../State';
import { insertDataAction } from './InsertData';
import { loadNextPageCard } from './LoadNextPage';

export function importDataAction(arg) {
    State.update(arg);
    State.setData(
        SKY.School.Lists.get(
            State.getList().id,
            SKY.ServiceManager.ResponseFormat.Array
        )
    );
    if (State.getData().length == SKY.PAGE_SIZE + 1 /* column labels */) {
        State.setPage(1);
        return Terse.CardService.replaceStack(loadNextPageCard());
    } else {
        return insertDataAction();
    }
}
global.action_lists_importData = importDataAction;
export default 'action_lists_importData';

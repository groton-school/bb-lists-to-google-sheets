import SKY from '../../SKY';
import State from '../../State';
import { insertDataAction } from './InsertData';
import { loadNextPageAction } from './LoadNextPage';

export function importDataAction(arg) {
    State.update(arg);
    State.setData(
        SKY.School.lists(State.getList().id, SKY.Options.ResponseFormat.Array)
    );
    if (State.getData().length == SKY.PAGE_SIZE + 1 /* column labels */) {
        State.setPage(1);
        return loadNextPageAction();
    } else {
        return insertDataAction();
    }
}
global.action_lists_importData = importDataAction;
export default 'action_lists_importData';

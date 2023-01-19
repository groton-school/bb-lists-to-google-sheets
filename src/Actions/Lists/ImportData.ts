import Lists from '../../Lists';
import School, { ResponseFormat } from '../../SKY/School';
import State from '../../State';
import { insertDataAction } from './InsertData';
import { loadNextPageAction } from './LoadNextPage';

export function importDataAction(arg) {
    State.update(arg);
    State.setData(School.lists(State.getList().id, ResponseFormat.Array));
    if (
        State.getData().length ==
        Lists.BLACKBAUD_PAGE_SIZE + 1 /* column labels */
    ) {
        State.setPage(1);
        return loadNextPageAction();
    } else {
        return insertDataAction();
    }
}
global.action_lists_importData = importDataAction;
export default 'action_lists_importData';

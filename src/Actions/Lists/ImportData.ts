import Lists from '../../Lists';
import SKY, { SkyResponse } from '../../SKY';
import State from '../../State';
import { insertDataAction } from './InsertData';
import { loadNextPageAction } from './LoadNextPage';

export function importDataAction(arg) {
    State.update(arg);
    State.setData(SKY.school.v1.lists(State.getList().id, SkyResponse.Array));
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
const ImportData = 'action_lists_importData';
export default ImportData;

import Sheets from '../../Sheets';
import { homeAction } from '../App/Home';

export function deleteMetadataAction() {
    Sheets.metadata.delete(Sheets.metadata.LIST);
    Sheets.metadata.delete(Sheets.metadata.RANGE);
    Sheets.metadata.delete(Sheets.metadata.NAME);
    Sheets.metadata.delete(Sheets.metadata.LAST_UPDATED);
    return homeAction();
}
global.action_sheets_deleteMetadata = deleteMetadataAction;
export default 'action_sheets_deleteMetadata';

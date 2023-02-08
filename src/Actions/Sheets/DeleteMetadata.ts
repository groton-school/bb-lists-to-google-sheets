import * as Sheets from '../../Sheets';
import { homeAction } from '../App/Home';

export function deleteMetadataAction() {
    Sheets.Metadata.removeList();
    Sheets.Metadata.removeRange();
    Sheets.Metadata.removeLastUpdated();
    return homeAction();
}

global.action_sheets_deleteMetadata = deleteMetadataAction;
export default 'action_sheets_deleteMetadata';

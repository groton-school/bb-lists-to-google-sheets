import { PREFIX } from './Constants';
import Sheets from './Sheets';
import State from './State';

export default class Lists {
    public static readonly UNCATEGORIZED = `${PREFIX}.Lists.uncategorized`;
    public static BLACKBAUD_PAGE_SIZE = 1000;

    public static setSheetName(sheet, timestamp = null) {
        sheet.setName(
            `${State.getList().name} (${timestamp || new Date().toLocaleString()})`
        );
        Sheets.metadata.set(Sheets.metadata.NAME, sheet.getName(), sheet);
    }
}

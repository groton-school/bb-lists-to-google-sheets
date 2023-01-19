import { PREFIX } from './Constants';
import State from './State';

export default class Sheets {
    public static metadata = class {
        public static readonly LIST = `${PREFIX}.list`;
        public static readonly RANGE = `${PREFIX}.range`;
        public static readonly LAST_UPDATED = `${PREFIX}.lastUpdated`;

        public static get(key, sheet = null) {
            sheet = sheet || State.getSheet();
            if (sheet) {
                const meta = sheet.createDeveloperMetadataFinder().withKey(key).find();
                if (meta && meta.length) {
                    const value = meta.shift().getValue();
                    try {
                        return JSON.parse(value);
                    } catch (e) {
                        return value;
                    }
                }
            }
            return null;
        }

        public static set(key, value, sheet = null) {
            sheet = sheet || State.getSheet();
            if (sheet) {
                const str = JSON.stringify(value);
                const meta = sheet.createDeveloperMetadataFinder().withKey(key).find();
                if (meta && meta.length) {
                    return meta.shift().setValue(str);
                } else {
                    return sheet.addDeveloperMetadata(key, str);
                }
            }
            return false;
        }

        public static delete(key, sheet = null) {
            sheet = sheet || State.getSheet();
            if (sheet) {
                const meta = sheet.createDeveloperMetadataFinder().withKey(key).find();
                if (meta && meta.length) {
                    return meta.shift().remove();
                }
            }
            return null;
        }
    };

    public static rangeToJSON(range) {
        return {
            row: range.getRow(),
            column: range.getColumn(),
            numRows: range.getNumRows(),
            numColumns: range.getNumColumns(),
            sheet: range.getSheet().getName(),
        };
    }

    public static rangeFromJSON(json) {
        // FIXME this fallback is unsafe without tracking tab changes!
        const sheet =
            State.getSpreadsheet().getSheetByName(json.sheet) || State.getSheet();
        return sheet.getRange(json.row, json.column, json.numRows, json.numColumns);
    }

    public static adjustRange(
        { row, column, numRows, numColumns },
        range = null,
        sheet = null
    ) {
        if (range) {
            sheet = range.getSheet();
            if (numRows > range.getNumRows()) {
                sheet.insertRows(range.getLastRow() + 1, numRows - range.getNumRows());
            }
            if (numColumns > range.getNumColumns()) {
                sheet.insertColumns(
                    range.getLastColumn() + 1,
                    numColumns - range.getNumColumns()
                );
            }
        } else if (sheet) {
            if (numRows < sheet.getMaxRows()) {
                sheet.deleteRows(numRows + 1, sheet.getMaxRows() - numRows);
            }
            if (numColumns < sheet.getMaxColumns()) {
                sheet.deleteColumns(numColumns + 1, sheet.getMaxColumns() - numColumns);
            }
        }
        return sheet.getRange(row, column, numRows, numColumns);
    }
}

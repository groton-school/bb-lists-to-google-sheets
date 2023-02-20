import * as g from '@battis/gas-lighter';
import * as SKY from '../SKY';
import ImportData from './ImportData';

const UNCATEGORIZED = `@@uncategorized@@`;

function groupCategories(
    categories: { [category: string]: SKY.School.Lists.Metadata[] },
    list: SKY.School.Lists.Metadata
) {
    if (list.id > 0) {
        if (!list.category) {
            list.category = UNCATEGORIZED;
        }
        if (!categories[list.category]) {
            categories[list.category] = [];
        }
        categories[list.category].push(list);
    }
    return categories;
}

export const getFunctionName = () => 'connect';
global.connect = () => {
    /**
     * FIXME detect if sheet or selection is already connected
     *   Route into Update or limit target options
     */
    SpreadsheetApp.getUi().showModalDialog(
        g.HtmlService.createTemplateFromFile('templates/connect', {
            thread: Utilities.getUuid(),
        }),
        'Connect'
    );
};

global.connectGetLists = () =>
    (SKY.School.Lists.get() as SKY.School.Lists.Metadata[]).reduce(
        groupCategories,
        {
            [UNCATEGORIZED]: [],
        }
    );

global.importData = ImportData;

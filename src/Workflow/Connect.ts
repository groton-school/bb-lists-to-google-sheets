import * as g from '@battis/gas-lighter';
import * as Metadata from '../Metadata';
import * as SKY from '../SKY';

const UNCATEGORIZED = 'Uncategorized';

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
    const prevList = Metadata.getList();
    const selection = SpreadsheetApp.getActive()
        .getActiveSheet()
        .getActiveRange()
        .getA1Notation();
    SpreadsheetApp.getUi().showModalDialog(
        g.HtmlService.createTemplateFromFile('templates/connect', {
            thread: Utilities.getUuid(),
            prevList,
            selection,
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

import App from './App';
import Lists from './Lists';
import Sheets from './Sheets';
import SKY from './SKY';
import State from './State';

const updateState = ({ parameters: { state } }) => {
    if (state) {
        state = JSON.parse(state);
        if (state.folder) {
            State.setFolder(DriveApp.getFolderById(state.folder));
        }

        var spreadsheet: GoogleAppsScript.Spreadsheet.Spreadsheet;
        if (state.spreadsheet) {
            spreadsheet = SpreadsheetApp.openById(state.spreadsheet);
        }
        var sheet: GoogleAppsScript.Spreadsheet.Sheet;
        if (state.sheet) {
            spreadsheet = spreadsheet || State.getSpreadsheet();
            sheet = spreadsheet.getSheetByName(state.sheet);
        }
        if (state.selection) {
            sheet = sheet || State.getSheet();
            State.setSelection(sheet.getRange(state.selection));
        } else if (sheet) {
            State.setSheet(sheet);
        } else if (spreadsheet) {
            State.setSpreadsheet(spreadsheet);
        }
        if (state.list) {
            State.setList(state.list);
        }
        if (state.page) {
            State.setPage(state.page);
        }
        if (state.data) {
            State.setData(state.data);
        }
        if (state.intent) {
            State.setIntent(state.intent);
        }
    }
};

global.__App_launch = (e) => {
    return App.launch(e);
};

global.__App_actions_home = (arg) => {
    updateState(arg);
    App.actions.home();
};

global.__Lists_actions_importData = (arg) => {
    updateState(arg);
    Lists.actions.importData();
};

global.__Lists_actions_lists = (arg) => {
    updateState(arg);
    Lists.actions.lists();
};

global.__Lists_actions_listDetail = (arg) => {
    updateState(arg);
    Lists.actions.listDetail();
};

global.__Lists_actions_loadNextPage = (arg) => {
    updateState(arg);
    Lists.actions.loadNextPage();
};

global.__Sheets_actions_deleteMetadata = (arg) => {
    updateState(arg);
    Sheets.actions.deleteMetadata();
};

global.__Sheets_actions_openSpreadsheet = (arg) => {
    updateState(arg);
    Sheets.actions.openSpreadsheet();
};

global.__Sheets_actions_showMetadata = (arg) => {
    updateState(arg);
    Sheets.actions.showMetadata();
};

global.__Sky_cardAuthorization = () => {
    return SKY.cardAuthorization();
};

global.__Sky_callbackAuthorization = (callbackRequest) => {
    return SKY.callbackAuthorization(callbackRequest);
};

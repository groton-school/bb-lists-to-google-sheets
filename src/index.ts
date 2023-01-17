import App from './App';
import Lists from './Lists';
import Sheets from './Sheets';
import SKY from './SKY';
import State from './State';

const updateState = ({ parameters: { state = null } }) => {
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

global.__App_launch = (event) => {
    return App.launch(event);
};

global.__App_actions_error = (arg) => {
    updateState(arg);
    return App.actions.error();
};

global.__App_actions_home = (arg) => {
    updateState(arg);
    return App.actions.home();
};

global.__Lists_actions_importData = (arg) => {
    updateState(arg);
    return Lists.actions.importData();
};

global.__Lists_actions_lists = (arg) => {
    updateState(arg);
    return Lists.actions.lists();
};

global.__Lists_actions_listDetail = (arg) => {
    updateState(arg);
    return Lists.actions.listDetail();
};

global.__Lists_actions_loadNextPage = (arg) => {
    updateState(arg);
    return Lists.actions.loadNextPage();
};

global.__Sheets_actions_breakConnection = (arg) => {
    updateState(arg);
    return Sheets.actions.breakConnection();
};

global.__Sheets_actions_deleteMetadata = (arg) => {
    updateState(arg);
    return Sheets.actions.deleteMetadata();
};

global.__Sheets_actions_openSpreadsheet = (arg) => {
    updateState(arg);
    return Sheets.actions.openSpreadsheet();
};

global.__Sheets_actions_showMetadata = (arg) => {
    updateState(arg);
    return Sheets.actions.showMetadata();
};

global.__Sky_cardAuthorization = () => {
    return SKY.cardAuthorization();
};

global.__Sky_callbackAuthorization = (callbackRequest) => {
    return SKY.callbackAuthorization(callbackRequest);
};

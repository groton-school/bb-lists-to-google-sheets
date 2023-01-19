import { Terse } from '@battis/google-apps-script-helpers';
import Home from '../App/Home';
import DeleteMetadata from './DeleteMetadata';

export function confirmBreakConnectionCard() {
    return Terse.CardService.newCard({
        header: `Are you sure?`,
        widgets: [
            `You are about to remove the developer metadata that connects this sheet to its Blackbaud data source. You will no longer be able to update the data on this sheet directly from Blackbaud. You will need to select the existing data and replace it with a new import from Blackbaud if you need to get new data.`,
            Terse.CardService.newTextButton({
                text: 'Delete Metadata',
                functionName: DeleteMetadata,
            }),
            Terse.CardService.newTextButton({ text: 'Cancel', functionName: Home }),
        ],
    });
}

export function confirmBreakConnectionAction() {
    return Terse.CardService.pushCard(confirmBreakConnectionCard());
}
global.action_sheets_confirmBreakConnection = confirmBreakConnectionAction;
export default 'action_sheets_confirmBreakConnection';

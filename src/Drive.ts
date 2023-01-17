export default class Drive {
    public static inferFolder(launchEvent) {
        if (
            launchEvent &&
            launchEvent.drive &&
            launchEvent.drive &&
            launchEvent.drive.activeCursorItem
        ) {
            const file = DriveApp.getFileById(launchEvent.drive.activeCursorItem.id);
            const parents = file.getParents();
            return parents.next() || null;
        }
        return null;
    }
}

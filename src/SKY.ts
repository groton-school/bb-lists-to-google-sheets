import SchoolClass, { Lists as ListsNamespace } from './SKY/School';
import ServiceManager, { Options } from './SKY/ServiceManager';

export namespace School {
    export namespace Lists {
        export type Metadata = ListsNamespace.Metadata;
        export namespace Data {
            export type List = ListsNamespace.Data.List;
        }
    }
}

export default class SKY {
    public static readonly PAGE_SIZE = 1000;
    public static ServiceManager = ServiceManager;
    public static School = SchoolClass;
    public static Options = Options;
}

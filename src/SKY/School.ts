import ServiceManager, { Options } from './ServiceManager';

export namespace Lists {
    export type Metadata = {
        id: number;
        name: string;
        type: 'Basic' | 'Advanced';
        description: string;
        category: string;
        created_by: string;
        created: string;
        last_modified: string;
    };
    export namespace Data {
        type Column = { name: string; value: string };
        type Row = { columns: Column[] };
        export type List = {
            count: number;
            page: number;
            results: { rows: Row[] };
        };
    }
}

const ResponseFormat = Options.ResponseFormat;

export default class School {
    public static lists(
        id?: number,
        format: Options.ResponseFormat = ResponseFormat.JSON,
        page?: number
    ) {
        if (id) {
            return School.listContent(id, format, page);
        } else {
            return School.listOfLists(format);
        }
    }

    private static listOfLists(format: Options.ResponseFormat) {
        const response = ServiceManager.makeRequest(
            'https://api.sky.blackbaud.com/school/v1/lists'
        ) as { count: number; value: Lists.Metadata[] };
        switch (format) {
            case ResponseFormat.JSON:
                return response.value;
            case ResponseFormat.Raw:
                return response;
        }
    }

    private static listContent(
        id: number,
        format: Options.ResponseFormat,
        page = 1
    ) {
        const response = ServiceManager.makeRequest(
            `https://api.sky.blackbaud.com/school/v1/lists/advanced/${id}?page=${page}`
        ) as Lists.Data.List;
        switch (format) {
            case ResponseFormat.JSON:
                return response.results.rows.map((row) =>
                    row.columns.reduce((obj, { name, value }) => {
                        obj[name] = value;
                        return obj;
                    }, {})
                );
            case ResponseFormat.Array:
                if (response.count == 0) {
                    return [];
                }
                const [first, ...rows] = response.results.rows;
                const arr = [
                    first.columns.reduce((arr, { name }): string[] => {
                        arr.push(name);
                        return arr;
                    }, []),
                    first.columns.reduce((arr, { value }): string[] => {
                        arr.push(value);
                        return arr;
                    }, []),
                ];
                rows.forEach((row) => {
                    arr.push(row.columns.map(({ value }) => value));
                });
                return arr;
            case ResponseFormat.Raw:
                return response;
        }
    }
}

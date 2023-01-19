import ServiceManager from './ServiceManager';

export enum ResponseFormat {
    JSON,
    Array,
    Raw,
}

export type ListMetadata = {
    id: number;
    name: string;
    type: 'Basic' | 'Advanced';
    description: string;
    category: string;
    created_by: string;
    created: string;
    last_modified: string;
};

type ColumnData = { name: string; value: string };
type RowData = { columns: ColumnData[] };
type ListData = { count: number; page: number; results: { rows: RowData[] } };

export default class School {
    public static lists(
        id?: string,
        format: ResponseFormat = ResponseFormat.JSON,
        page?: number
    ) {
        if (id) {
            return School.listContent(id, format, page);
        } else {
            return School.listOfLists(format);
        }
    }

    private static listOfLists(format: ResponseFormat) {
        const response = ServiceManager.makeRequest(
            'https://api.sky.blackbaud.com/school/v1/lists'
        ) as { count: number; value: ListMetadata[] };
        switch (format) {
            case ResponseFormat.JSON:
                return response.value;
            case ResponseFormat.Raw:
                return response;
        }
    }

    private static listContent(id: string, format: ResponseFormat, page = 1) {
        const response = ServiceManager.makeRequest(
            `https://api.sky.blackbaud.com/school/v1/lists/advanced/${id}?page=${page}`
        ) as ListData;
        switch (format) {
            case ResponseFormat.JSON:
                return response.results.rows.map((row) =>
                    row.columns.reduce((obj, { name, value }) => {
                        obj[name] = value;
                        return obj;
                    }, {})
                );
            case ResponseFormat.Array:
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

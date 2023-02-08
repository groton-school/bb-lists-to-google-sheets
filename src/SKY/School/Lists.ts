import * as ServiceManager from '../ServiceManager';

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

const ResponseFormat = ServiceManager.ResponseFormat;

export function get(
    id?: number,
    format: ServiceManager.ResponseFormat = ServiceManager.ResponseFormat.JSON,
    page?: number
) {
    if (id) {
        return listContent(id, format, page);
    } else {
        return listOfLists(format);
    }
}

export function listOfLists(format: ServiceManager.ResponseFormat) {
    const response = ServiceManager.makeRequest(
        'https://api.sky.blackbaud.com/school/v1/lists'
    ) as { count: number; value: Metadata[] };
    switch (format) {
        case ResponseFormat.JSON:
            return response.value;
        case ResponseFormat.Raw:
            return response;
    }
}

function listContent(
    id: number,
    format: ServiceManager.ResponseFormat,
    page = 1
) {
    const response = ServiceManager.makeRequest(
        `https://api.sky.blackbaud.com/school/v1/lists/advanced/${id}?page=${page}`
    ) as Data.List;
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

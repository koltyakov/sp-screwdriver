export interface ISetItemsProperty {
    id?: number;
    field: string;
    value: string;
}

export interface ISetItemsPropertiesParams {
    baseUrl?: string;
    listPath: string;
    itemId: number;
    properties: ISetItemsProperty[];
    updateId?: number;
    queryId?: number;
}

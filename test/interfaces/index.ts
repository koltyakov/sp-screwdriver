import { INewTerm } from './../../src/interfaces/IMmd';

export interface IEnvironmentConfig {
    environmentName: string;
    authConfigPath: string;
    paramsConfigPath: string;
}

export interface IUpsConfig {
    accountName: string;
}

export interface IMmdConfig {
    sspId: string;
    serviceName: string;
    termSetId: string;
    termId: string;
    lcid?: number;
    newTerms: INewTerm[];
}

export interface IVersionsDocumentsConfig {
    fileName: string;
}

export interface IVersionsItemsConfig {
    listId: string;
    itemId: number;
    fieldName: string;
}

export interface IVersionsConfig {
    documents: IVersionsDocumentsConfig;
    items: IVersionsItemsConfig;
}

export interface IItemsConfig {
    listPath: string;
    itemId: number;
    properties: ISetItemsProperty[];
}

export interface ITestConfig {
    versions: IVersionsConfig;
    items: IItemsConfig;
    mmd: IMmdConfig;
    ups: IUpsConfig;
}

export interface ISetItemsProperty {
    field: string;
    value: string;
}

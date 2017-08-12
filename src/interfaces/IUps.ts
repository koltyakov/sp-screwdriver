export interface IGetUserProfileByNameProperties {
    baseUrl?: string;
    accountName: string;
}

export interface INewPropData {
    name: string;
    values: string[];

    privacy?: 'NotSet' | string;
    isPrivacyChanged?: boolean;
    isValueChanged?: boolean;
}

export interface IModifyUserPropertyByAccountNameProperties {
    baseUrl?: string;
    accountName: string;
    newData: INewPropData[];
}

export interface IGetUserPropertyByAccountNameProperties {
    baseUrl?: string;
    accountName: string;
    propertyName: string;
}

export interface IGetUserProfilePropertyForProperties {
    baseUrl?: string;
    accountName: string;
    propertyName: string;
}

export interface IGetPropertiesForProperties {
    baseUrl?: string;
    accountName: string;
}

export interface ISetSingleValueProfilePropertyProperties {
    baseUrl?: string;
    accountName: string;
    propertyName: string;
    propertyValue: string;
}

export interface ISetMultiValuedProfilePropertyProperties {
    baseUrl?: string;
    accountName: string;
    propertyName: string;
    propertyValues: string[];
}

export interface IUserProp {
    name: string;
    values: any[];
    privacy: string;
    isPrivacyChanged: boolean;
    isValueChanged: boolean;
}

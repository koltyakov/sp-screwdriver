export interface IGetTermSetsParams {
    baseUrl?: string;
    sspId: string;
    termSetId: string;

    lcid?: number;
    clientTimeStamp?: string;
    clientVersion?: number;
}

export interface ITermSet {
    name: string;
    id: string;
    _raw: any;
}

export interface ITerm {
    name: string;
    id: string;
    enableForTagging: boolean;
    parentId: string;
    termSetId: string;
    _raw: any;
}

export interface ITermSetsResponse {
    termSet: ITermSet;
    terms: ITerm[];
}

export interface IGetChildTermsInTermSetParams {
    baseUrl?: string;
    sspId: string;
    termSetId: string;

    lcid?: number;
}

export interface IGetChildTermsInTermParams {
    baseUrl?: string;
    sspId: string;
    termId: string;
    termSetId: string;

    lcid?: number;
}

export interface IGetTermsByLabelParams {
    baseUrl?: string;
    label: string;
    termIds: string[];

    matchOption?: 'ExactMatch' | 'StartsWith';
    resultCollectionSize?: number;
    addIfNotFound?: boolean;
    lcid?: number;
}

export interface IGetKeywordTermsByGuidsParams {
    baseUrl?: string;
    termIds: string[];

    lcid?: number;
}

export interface INewTerm {
    label: string;
    parentTermId?: string;
}

export interface IAddTermsParams {
    baseUrl?: string;
    sspId: string;
    termSetId: string;
    newTerms: INewTerm[];

    lcid?: number;
}

export interface IGetAllTermsParams {
    baseUrl?: string;
    serviceName: string;
    termSetId: string;
    properties?: string[];
}

export interface ISetTermNameParams {
    baseUrl?: string;
    serviceName: string;
    termSetId: string;
    termId: string;
    newName: string;
}

export interface IDeprecateTermsParams {
    baseUrl?: string;
    serviceName: string;
    termSetId: string;
    termId: string;
    deprecate?: boolean;
}

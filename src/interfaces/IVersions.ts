export interface IDocumentVersions {
  baseUrl?: string;
  fileName: string;
}

export interface IDocumentVersion extends IDocumentVersions {
  fileVersion: string;
}

export interface IItemVersions {
  baseUrl?: string;
  listId: string;
  itemId: number;
  fieldName: string;
}

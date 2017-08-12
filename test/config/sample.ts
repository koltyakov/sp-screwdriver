import { ITestConfig } from './../interfaces';

export const testConfigs: ITestConfig = {
    items: {
        listPath: '/_catalogs/masterpage', // Can be relative
        itemId: 3746,
        properties: [{
            field: 'PublishingAssociatedContentType',
            value: ';#Article Page;#0x010100C568DB52D9D0A14D9B2FDCC96666E9F2007948130EC3DB064584E219954237AF3900242457EFB8B24247815D688C526CD44D;#'
        }]
    },
    mmd: {
        sspId: 'ffca190d3d644312ad74dc2c05fd27fc',
        serviceName: 'Taxonomy_5KSgChEZ9j15+7UVInQNRQ==',
        termSetId: '8ed8c9ea-7052-4c1d-a4d7-b9c10bffea6f',
        termId: 'e66280de-4fdd-4cb9-8783-ce3efe3f7ef8',
        lcid: 1033,
        newTerms: [{
            label: 'Dev2',
            parentTermId: '533b673d-85f3-4654-a8be-74777992adba'
        }]
    },
    ups: {
        accountName: 'i:0#.f|membership|username'
    },
    versions: {
        documents: {
            fileName: 'Shared%20Documents/Filename.xlsx'
        },
        items: {
            listId: 'BE710A75-3CF7-4D62-BB95-C317FFAA7905',
            itemId: 1,
            fieldName: 'Title'
        }
    }
};

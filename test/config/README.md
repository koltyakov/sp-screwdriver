
## Items

### Property bag `PublishingAssociatedContentType`

#### Sample

```javascript
export const testConfigs: ITestConfig = {
    ...
    items: {
        listPath: '/_catalogs/masterpage',
        itemId: 3746,
        properties: [{
            field: 'PublishingAssociatedContentType',
            value: ';#Article Page;#0x010100C568DB52D9D0A14D9B2FDCC96666E9F2007948130EC3DB064584E219954237AF3900242457EFB8B24247815D688C526CD44D;#'
        }]
    }
    ...
};
```

#### How to get a value for testing

- Go to `/_catalogs/masterpage`
- Create Page Layout
- Check for PublishingAssociatedContentType

```javascript
import * as pnp from 'pnp';

let itemId = 2502; // change on your own

pnp.sp.web
    .getList(`${_spPageContextInfo.webServerRelativeUrl}/_catalogs/masterpage`)
    .items.getById(itemId)
    .select('Properties/PublishingAssociatedContentType')
    .expand('Properties').get()
    .then(item => {
        console.log(item.Properties.PublishingAssociatedContentType);
    })
    .catch(console.log);
```

- Change PublishingAssociatedContentType to a random one
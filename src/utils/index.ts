import { parseString } from 'xml2js';

export class Utils {

    public trimMultiline = (multiline) => {
        return multiline.split('\n').map(line => line.trim()).join('');
    }

    public parseXml = (xmlString: string): Promise<any> => {
        return new Promise((resolve, reject) => {
            parseString(xmlString, (err, result) => {
                if (err) {
                    return reject(err);
                }
                resolve(result);
            });
        });
    }

    public soapHeaders = (soapBody: string): Headers => {
        let headers: Headers = new Headers();

        headers.set('Accept', 'application/xml, text/xml, */*; q=0.01');
        headers.set('Content-Type', 'text/xml;charset="UTF-8"');
        headers.set('X-Requested-With', 'XMLHttpRequest');
        headers.set('Content-Length', soapBody.length.toString());

        return headers;
    }

    public csomHeaders = (requestBody: string, digest: string): Headers => {
        let headers: Headers = new Headers();

        headers.set('Accept', '*/*');
        headers.set('Content-Type', 'text/xml;charset="UTF-8"');
        headers.set('X-Requested-With', 'XMLHttpRequest');
        headers.set('Content-Length', requestBody.length.toString());
        headers.set('X-RequestDigest', digest);

        return headers;
    }

}

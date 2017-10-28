import { parseString } from 'xml2js';

export class Utils {

  public soapEnvelope = (body: string): string => {
    const envelopMeta: string = 'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" ' +
      'xmlns:xsd="http://www.w3.org/2001/XMLSchema" ' +
      'xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"';

    return this.trimMultiline(`
            <?xml version="1.0" encoding="utf-8"?>
            <soap:Envelope ${envelopMeta}>
                <soap:Body>
                    ${body}
                </soap:Body>
            </soap:Envelope>
        `);
  }

  public trimMultiline = (multiline) => {
    return multiline
      .split('\n')
      .map(line => line.trim())
      .filter(line => line.length > 0)
      .join('').trim();
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

  public soapHeaders = (soapBody: string): any => {
    return {
      'Accept': 'application/xml, text/xml, */*; q=0.01',
      'Content-Type': 'text/xml;charset="UTF-8"',
      'X-Requested-With': 'XMLHttpRequest',
      'Content-Length': soapBody.length
    };
  }

  public csomHeaders = (requestBody: string, digest: string): any => {
    return {
      'Accept': '*/*',
      'Content-Type': 'text/xml;charset="UTF-8"',
      'X-Requested-With': 'XMLHttpRequest',
      'Content-Length': requestBody.length,
      'X-RequestDigest': digest
    };
  }

  public relativeFromAbsoluteUrl = (absoluteUrl: string): string => {
    return `/${absoluteUrl.replace('://', '').split('/').splice(1, 100).join('/')}`;
  }

  public toInnerXmlPackage = (xml: string): string => {
    return xml.replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }

}

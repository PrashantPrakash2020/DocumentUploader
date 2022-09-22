import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
export class DataService {
    private context: WebPartContext;
    private absoluteURL: string;
    constructor(context: WebPartContext) {
        this.context = context;
        this.absoluteURL = context.pageContext.web.absoluteUrl;
    }
    public updateListItemById(listName: string, id: string, body: string): Promise<SPHttpClientResponse> {
        const restUrl: string = this.absoluteURL + `/_api/web/lists/GetByTitle('${listName}')/items(${id})`;
        return this.context.spHttpClient.post(restUrl,
            SPHttpClient.configurations.v1, {
            headers: {
                'Accept': 'application/json;odata=verbose',
                'Content-type': 'application/json;odata=verbose',
                'odata-version': '',
                'IF-MATCH': '*',
                'X-HTTP-Method': 'MERGE'
            },
            body: body
        });
    }
    public uploaddocument(listName: string, fileName: string,
        file: File): Promise<SPHttpClientResponse> {
        return this.context.spHttpClient.post(this.context.pageContext.web.absoluteUrl +
            `/_api/web/lists/getbytitle('${listName}')/RootFolder/Files/Add(url='${fileName}', overwrite=true)`,
            SPHttpClient.configurations.v1,
            {
                headers: {
                    'Accept': 'application/json',
                    'Content-type': 'application/json'
                },
                body: file
            });
    }
    public getDocId(listName: string, filename: string): Promise<SPHttpClientResponse> {
        const restURL = this.context.pageContext.web.absoluteUrl +
            `/_api/web/lists/GetByTitle('${listName}')/items?$filter=FileLeafRef eq '${filename}'&$select=Id`;
        return this.context.spHttpClient.get(restURL,
            SPHttpClient.configurations.v1).then((data: SPHttpClientResponse) => {
                return data;
            });
    }
    public getchoice(listName: string, columnname: string): Promise<SPHttpClientResponse> {
        const restURL = this.context.pageContext.web.absoluteUrl +
            `/_api/web/lists/GetByTitle('${listName}')/fields?$filter=EntityPropertyName eq '${columnname}'`;
        return this.context.spHttpClient.get(restURL,
            SPHttpClient.configurations.v1).then((data: SPHttpClientResponse) => {
                return data;
            });
    }
}
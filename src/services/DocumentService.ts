import { SPHttpClient,
    SPHttpClientResponse,
    ISPHttpClientOptions
} from "@microsoft/sp-http";

import { ILibraryDocument } from '../models'

// import {
//     DetailsList,
//     DetailsListLayoutMode,
//     Selection,
//     SelectionMode,
//     IColumn
//   } from 'office-ui-fabric-react/lib/DetailsList';


export class DocumentService {
private _spHttpOptions: any = {
    getNoMetadata: <ISPHttpClientOptions> {
        headers: { 'ACCEPT': 'application/json; odata.metadata=none' }
    }
}

constructor(private siteAbsoluteUrl: string, private client: SPHttpClient) {}

getDocuments(listName: string, queryString: string): Promise<ILibraryDocument[]>{
    const LIST_API_ENPOINT: string = `/_api/web/lists/getbytitle('${listName}')`;
    // const SELECT_QUERY: string =  `$select=description0,someColumn,File_x0020_Type,Title,Modified,Created,Author/Title,Editor/Title&$expand=Editor,Author,File`
    // const FILTER_QUERY: string = '$filter=OData__ModerationStatus eq 0'

    //EDIT: In SharePoint Online the best way to achieve a url that opens a document in browser editing ‘web=1’ as a query string parameter to the file URL. E.g. https://tenant.sharepoint.com/Documents/file.docx?web=1

    let promise: Promise<ILibraryDocument[]> = new Promise<ILibraryDocument[]>((resolve, reject) => {
        let query = `${this.siteAbsoluteUrl}${LIST_API_ENPOINT}/Items?${queryString}`

        this.client.get(
            query,
            SPHttpClient.configurations.v1,
            this._spHttpOptions.getNoMetadata
        )
        .then((response: SPHttpClientResponse): Promise< { value: ILibraryDocument[] } > => {
            return response.json()
        })
        .then((response: { value: ILibraryDocument[] }) => {
            resolve(response.value);
        })
        .catch((error:any) => {
            reject(error);
        });
    });

    return promise;
}

/**
 * PRIVATE
 * Flatten a deep object into a one level object with it’s path as key
 *
 * @param  {object} object - The object to be flattened
 *
 * @return {object}        - The resulting flat object
 */
flattenObject(object, separator = '/') {
    return (<any>Object).assign({}, ...function _flatten(child, path = []) {
        return [].concat(...Object.keys(child).map(key => (child[key] !== null &&  typeof child[key] === 'object')
            ? _flatten(child[key], path.concat([key]))
            : ({ [path.concat([key]).join(separator)] : child[key] })
        ));
    }(object));
}


calcFileSize(length: number) {
    let i = Math.floor( Math.log(length) / Math.log(1024) );
    return (+( length / Math.pow(1024, i) ).toFixed(1) * 1).toString() + ' ' + ['B', 'KB', 'MB', 'GB', 'TB'][i];
}

sortDocuments(items: Array<ILibraryDocument>, sortBy: string, direction = 'desc'): Array<ILibraryDocument>{
    //check if sortby property is not an object, can't sort if true
    if(typeof(items[0][sortBy]) == 'object') {
        return items
    }
   
    if (direction == 'desc') {
        return items.sort((a: ILibraryDocument, b: ILibraryDocument) => {
            let _valuea, _valueb;
            if(typeof(a[sortBy])==='string') {
                _valuea = a[sortBy].toLowerCase()
                _valueb = b[sortBy].toLowerCase()
            } else {
                _valuea = a[sortBy];
                _valueb = b[sortBy];
            }
    
            if (_valuea < _valueb) {
                return 1;
            }
            if (_valuea > _valueb) {
                return -1;
            }
            return 0;
        });
    } else {
        return items.sort((a: ILibraryDocument, b: ILibraryDocument) => {
            let _valuea, _valueb;

            if(typeof(a[sortBy])==='string') {
                _valuea = a[sortBy].toLowerCase()
                _valueb = b[sortBy].toLowerCase()
            } else {
                _valuea = a[sortBy];
                _valueb = b[sortBy];
            }

            if (_valuea < _valueb) {
                return -1;
            }
            if (_valuea > _valueb) {
                return 1;
            }
            return 0;
            });
    }
}

limitDocuments(results: Array<ILibraryDocument>, limitNo: number): Array<ILibraryDocument> {
    let _resultArray;

    typeof(limitNo) == 'number' 
    ? _resultArray = results.slice(0, (limitNo))
    : _resultArray = results
    
    return _resultArray
}



}
import { IColumn } from 'office-ui-fabric-react/lib/DetailsList';

export interface ITableState {
    columns: IColumn[];
    items: IDocument[];
    selectionDetails: string;
    isModalSelection: boolean;
    isCompactMode: boolean;
  }
  
export interface IDocument {
    [key: string]: any;
    name: string;
    value: string;
    iconName: string;
    modifiedBy: string;
    dateModified: string;
    dateModifiedValue: number;
    fileSize: string;
    fileSizeRaw: number;
    url: string;
  }
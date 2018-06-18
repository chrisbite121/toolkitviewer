// import * as React from 'react';
// import styles from '../HelloWorld.module.scss';
import { ITableProps } from "./ITableProps";
// import { escape } from '@microsoft/sp-lodash-subset';

/* tslint:disable:no-unused-variable */
import * as React from "react";
/* tslint:enable:no-unused-variable */
import { TextField } from "office-ui-fabric-react/lib/TextField";
import { Toggle } from "office-ui-fabric-react/lib/Toggle";
import {
  DetailsList,
  DetailsListLayoutMode,
  Selection,
  SelectionMode,
  IColumn
} from "office-ui-fabric-react/lib/DetailsList";
import { MarqueeSelection } from "office-ui-fabric-react/lib/MarqueeSelection";
import { lorem } from "@uifabric/example-app-base";
// import './DetailsListExample.scss';

import { IDocument, ITableState } from "../../../../models";

import { TableService } from "../../../../services";

let _items: IDocument[] = [];

const fileIcons: { name: string }[] = [
  { name: "accdb" },
  { name: "csv" },
  { name: "docx" },
  { name: "dotx" },
  { name: "mpp" },
  { name: "mpt" },
  { name: "odp" },
  { name: "ods" },
  { name: "odt" },
  { name: "one" },
  { name: "onepkg" },
  { name: "onetoc" },
  { name: "potx" },
  { name: "ppsx" },
  { name: "pptx" },
  { name: "pub" },
  { name: "vsdx" },
  { name: "vssx" },
  { name: "vstx" },
  { name: "xls" },
  { name: "xlsx" },
  { name: "xltx" },
  { name: "xsn" }
];

export class DetailsListDocuments extends React.Component<
  ITableProps,
  ITableState
> {
  private _selection: Selection;
  private tableService: TableService;

  constructor(props: any) {
    super(props);

    this.tableService = new TableService();

    const _columns: IColumn[] = this._getColumns();

    this._selection = new Selection({
      onSelectionChanged: () => {
        this.setState({
          selectionDetails: this._getSelectionDetails(),
          isModalSelection: this._selection.isModal()
        });
      }
    });

    this.state = {
      items: _items,
      columns: _columns,
      selectionDetails: this._getSelectionDetails(),
      isModalSelection: this._selection.isModal(),
      isCompactMode: true
    };
  }

  componentWillReceiveProps(newProps) {
    const _columns: IColumn[] = this._getColumns();

    _items = [];
    newProps.data.forEach(file => {
      let date = { value: null, dateFormatted: null };
      if (file.hasOwnProperty("Modified")) {
        date = this._generateDate(new Date(file.Modified));
      }

      let _urlValue;
      if (
        file.hasOwnProperty("File/LinkingUri") &&
        file["File/LinkingUri"] !== null
      ) {
        _urlValue = decodeURIComponent(file["File/LinkingUri"]);
      } else if (file.hasOwnProperty("File/ServerRelativeUrl")) {
        _urlValue =
          file["BaseUrl"].split("/sites/")[0] + file["File/ServerRelativeUrl"];
      } else {
        _urlValue = "";
      }

      let _nameValue, _value;
      //add guard logic for each column
      if (file.hasOwnProperty("File/Name") && file["File/Name"] !== null) {
        //Remove trailing file extension (e.g. .docx) from filename if exists
        if (
          file["File/Name"].substring(0, file["File/Name"].lastIndexOf(".")) !==
          -1
        ) {
          _nameValue = file["File/Name"].substring(
            0,
            file["File/Name"].lastIndexOf(".")
          );
          _value = file["File/Name"].substring(
            0,
            file["File/Name"].lastIndexOf(".")
          );
        } else {
          _nameValue = file["File/Name"];
          _value = file["File/Name"];
        }
      } else if (file.hasOwnProperty("Title")) {
        _nameValue = file.Title;
        _value = file.Title;
      } else {
        _nameValue = "<BLANK>";
        _value = "<BLANK>";
      }

      let _fileSize, _fileSizeRaw;
      if (file.hasOwnProperty("Size") && file.hasOwnProperty("File/Length")) {
        _fileSize = file.Size;
        _fileSizeRaw = file["File/Length"];
      } else {
        _fileSize = "";
        _fileSizeRaw = "";
      }

      //version number
      let _version;
      if (file.hasOwnProperty("Versions/0/VersionLabel")) {
        _version = file["Versions/0/VersionLabel"];
      } else if (file.hasOwnProperty("FullVersion")) {
        _version = file["FullVersion"];
      } else {
        _version = "";
      }

      const fileType = this._generateFileIcon(file.File_x0020_Type);
      _items.push({
        name: _nameValue,
        value: _value,
        iconName: fileType.url,
        modifiedBy: file["Editor/Title"],
        dateModified: date.dateFormatted,
        dateModifiedValue: date.value,
        fileSize: _fileSize,
        fileSizeRaw: _fileSizeRaw,
        url: _urlValue,
        version: _version
      });
    });

    console.error(newProps);

    this.setState({
      items: _items,
      columns: _columns,
      selectionDetails: this._getSelectionDetails(),
      isModalSelection: this._selection.isModal(),
      isCompactMode: true
    });
  }

  public render() {
    const { columns, isCompactMode, items, selectionDetails } = this.state;

    return (
      <div>
        <MarqueeSelection selection={this._selection}>
          <DetailsList
            items={items}
            compact={isCompactMode}
            columns={columns}
            selectionMode={
              this.state.isModalSelection
                ? SelectionMode.multiple
                : SelectionMode.none
            }
            setKey="set"
            layoutMode={DetailsListLayoutMode.justified}
            isHeaderVisible={true}
            selection={this._selection}
            selectionPreservedOnEmptyClick={true}
            onItemInvoked={this._onItemInvoked}
            enterModalSelectionOnTouch={true}
          />
        </MarqueeSelection>
      </div>
    );
  }

  public componentDidUpdate(previousProps: any, previousState: ITableState) {
    if (previousState.isModalSelection !== this.state.isModalSelection) {
      this._selection.setModal(this.state.isModalSelection);
    }
  }

  private _onChangeCompactMode = (checked: boolean): void => {
    this.setState({ isCompactMode: checked });
  };

  private _onChangeModalSelection = (checked: boolean): void => {
    this.setState({ isModalSelection: checked });
  };

  private _onChangeText = (text: any): void => {
    this.setState({
      items: text
        ? _items.filter(i => i.name.toLowerCase().indexOf(text) > -1)
        : _items
    });
  };

  private _onItemInvoked(item: any): void {
    alert(`Item invoked: ${item.name}`);
  }

  private _generateDate(date: Date): { value: number; dateFormatted: string } {
    const dateData = {
      value: date.valueOf(),
      dateFormatted: date.toLocaleDateString("en-GB")
    };
    return dateData;
  }

  private _generateFileIcon(docType: string): { docType: string; url: string } {
    // const docType: string = fileIcons[Math.floor(Math.random() * fileIcons.length) + 0].name;

    let _match = null;
    if (docType) {
      fileIcons.forEach(icon => {
        icon.name == docType ? (_match = true) : null;
      });
    }

    if (!docType || docType == null || !_match) {
      return {
        docType,
        url:
          "https://static2.sharepointonline.com/files/fabric/assets/brand-icons/product/svg/sharepoint_16x1.svg"
      };
    } else {
      return {
        docType,
        url: `https://static2.sharepointonline.com/files/fabric/assets/brand-icons/document/svg/${docType}_16x1.svg`
      };
    }
  }

  private _getSelectionDetails(): string {
    const selectionCount = this._selection.getSelectedCount();

    switch (selectionCount) {
      case 0:
        return "No items selected";
      case 1:
        return (
          "1 item selected: " + (this._selection.getSelection()[0] as any).name
        );
      default:
        return `${selectionCount} items selected`;
    }
  }

  private _onColumnClick = (
    ev: React.MouseEvent<HTMLElement>,
    column: IColumn
  ): void => {
    const { columns, items } = this.state;
    let newItems: IDocument[] = items.slice();
    const newColumns: IColumn[] = columns.slice();
    const currColumn: IColumn = newColumns.filter(
      (currCol: IColumn, idx: number) => {
        return column.key === currCol.key;
      }
    )[0];
    newColumns.forEach((newCol: IColumn) => {
      if (newCol === currColumn) {
        currColumn.isSortedDescending = !currColumn.isSortedDescending;
        currColumn.isSorted = true;
      } else {
        newCol.isSorted = false;
        newCol.isSortedDescending = true;
      }
    });
    newItems = this._sortItems(
      newItems,
      currColumn.fieldName,
      currColumn.isSortedDescending
    );
    this.setState({
      columns: newColumns,
      items: newItems
    });
  };

  private _sortItems = (
    items: IDocument[],
    sortBy: string,
    descending = false
  ): IDocument[] => {
    if (descending) {
      return items.sort((a: IDocument, b: IDocument) => {
        let _valuea, _valueb;

        if (typeof a[sortBy] === "string" && typeof b[sortBy] === "string") {
          _valuea = a[sortBy].toLowerCase();
          _valueb = b[sortBy].toLowerCase();
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
      return items.sort((a: IDocument, b: IDocument) => {
        let _valuea, _valueb;

        if (typeof a[sortBy] === "string" && typeof b[sortBy] === "string") {
          _valuea = a[sortBy].toLowerCase();
          _valueb = b[sortBy].toLowerCase();
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
  };

  private _getColumns(): IColumn[] {
    return [
      {
        key: "column1",
        name: "File Type",
        headerClassName: "DetailsListExample-header--FileIcon",
        className: "DetailsListExample-cell--FileIcon",
        iconClassName: "DetailsListExample-Header-FileTypeIcon",
        iconName: "Page",
        isIconOnly: true,
        fieldName: "name",
        minWidth: 10,
        maxWidth: 16,
        onRender: (item: IDocument) => {
          return <img src={item.iconName} />;
        }
      },
      {
        key: "column2",
        name: "Name",
        fieldName: "name",
        minWidth: 150,
        maxWidth: 200,
        isRowHeader: true,
        isResizable: true,
        // isSorted: true,
        // isSortedDescending: false,
        onColumnClick: this._onColumnClick,
        data: "string",
        isPadded: true,
        onRender: (item: IDocument) => {
          return (
            <a className="ms-listlink ms-draggable" href={item.url}>
              {item.name}
            </a>
          );
        }
      },
      {
        key: "column3",
        name: "Date Modified",
        fieldName: "dateModifiedValue",
        minWidth: 40,
        maxWidth: 70,
        isResizable: true,
        onColumnClick: this._onColumnClick,
        data: "number",
        onRender: (item: IDocument) => {
          return <span>{item.dateModified}</span>;
        },
        isPadded: true
      },
      {
        key: "column4",
        name: "Modified By",
        fieldName: "modifiedBy",
        minWidth: 40,
        maxWidth: 90,
        isResizable: true,
        isCollapsable: true,
        data: "date",
        onColumnClick: this._onColumnClick,
        onRender: (item: IDocument) => {
          return <span>{item.modifiedBy}</span>;
        },
        isPadded: true
      },
      // {
      //   key: 'column5',
      //   name: 'File Size',
      //   fieldName: 'fileSizeRaw',
      //   minWidth: 70,
      //   maxWidth: 90,
      //   isResizable: true,
      //   isCollapsable: true,
      //   data: 'number',
      //   onColumnClick: this._onColumnClick,
      //   onRender: (item: IDocument) => {
      //     return (
      //       <span>
      //         { item.fileSize }
      //       </span>
      //     );
      //   }
      // },
      {
        key: "column5",
        name: "Version",
        fieldName: "version",
        minWidth: 40,
        maxWidth: 70,
        isResizable: true,
        isCollapsable: true,
        data: "string",
        onColumnClick: this._onColumnClick,
        onRender: (item: IDocument) => {
          return <span>{item.version}</span>;
        }
      }
    ];
  }
}

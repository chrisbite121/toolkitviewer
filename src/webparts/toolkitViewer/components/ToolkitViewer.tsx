import * as React from "react";
import styles from "./ToolkitViewer.module.scss";

import { IToolkitViewerProps } from "./IToolkitViewerProps";
import { IToolkitViewerState } from "./IToolkitViewerState";

import { escape } from "@microsoft/sp-lodash-subset";

import { ILibraryDocument } from "../../../models";
import { DocumentService } from "../../../services";

import { DetailsListDocuments } from "./Table/Table";

export default class ToolkitViewer extends React.Component<
  IToolkitViewerProps,
  IToolkitViewerState
> {
  private documentService: DocumentService;

  constructor(props) {
    super(props);

    //init state
    this.state = { data: [] };

    this._getDocuments = this._getDocuments.bind(this);
  }

  componentWillUpdate(nextProps, nextState) {
    console.log("updating component");
  }

  componentDidMount() {
    this.documentService = new DocumentService(
      this.props.context.pageContext.web.absoluteUrl,
      this.props.context.spHttpClient
    );

    this._getDocuments(null, this.props);
  }

  componentWillReceiveProps(nextProps, nextState) {
    console.log("component updated");
    console.log(this.props);
    this._getDocuments(null, nextProps);
  }

  private _getDocuments(event, properties): void {
    if (event && event !== null) {
      try {
        event.preventDefault();
      } catch (error) {
        console.error(error);
      }
    }

    let promiseArray = this._constructPromiseArray(properties);

    if (promiseArray.length > 0) {
      Promise.all(promiseArray).then((resultArray: ILibraryDocument[][]) => {
        let _array = [];
        if (
          resultArray &&
          Array.isArray(resultArray) &&
          resultArray.length > 0
        ) {
          resultArray.forEach((result: ILibraryDocument[]) => {
            if (Array.isArray(result) && result.length > 0) {
              _array = _array.concat(result);
            }
          });
        }

        console.error(_array);

        //flatten results
        let _flattenArray = _array.map(item =>
          this.documentService.flattenObject(item)
        );

        _flattenArray.forEach(item => {
          //add file size property
          if (item.hasOwnProperty("File/Length")) {
            item["Size"] = this.documentService.calcFileSize(
              +item["File/Length"]
            );
          }
          //add full version
          if (
            item.hasOwnProperty("File/MajorVersion") &&
            item.hasOwnProperty("File/MinorVersion")
          ) {
            item["FullVersion"] = `${item["File/MajorVersion"]}.${
              item["File/MinorVersion"]
            }`;
          }

          //add base url
          item["BaseUrl"] = this.props.context.pageContext.web.absoluteUrl;
        });

        console.log("Flattened Result Objects");
        console.error(_flattenArray);

        let _sortedArray;
        if (
          this.props.orderBy &&
          this.props.direction &&
          this.props.itemLimit
        ) {
          _sortedArray = this.documentService.sortDocuments(
            _flattenArray,
            this.props.orderBy,
            this.props.direction
          );
          _sortedArray = this.documentService.limitDocuments(
            _sortedArray,
            +this.props.itemLimit
          );
        } else {
          _sortedArray = _array;
        }

        if (_array.length > 0) {
          this.setState({ data: _sortedArray });
        }
      });
    } else {
      console.error("no library definitions found, skipping get data call");
    }
  }

  _constructPromiseArray(properties: IToolkitViewerProps): Array<any> {
    let _array = [];

    //check if property value is specified before tyring make api call
    if (properties.library1 && properties.library1.length > 0) {
      _array.push(
        this.documentService.getDocuments(
          properties.library1,
          properties.queryString
        )
      );
    }

    if (properties.library2 && properties.library2.length > 0) {
      _array.push(
        this.documentService.getDocuments(
          properties.library2,
          properties.queryString
        )
      );
    }

    if (properties.library3 && properties.library3.length > 0) {
      _array.push(
        this.documentService.getDocuments(
          properties.library3,
          properties.queryString
        )
      );
    }

    if (properties.library4 && properties.library4.length > 0) {
      _array.push(
        this.documentService.getDocuments(
          properties.library4,
          properties.queryString
        )
      );
    }
    return _array;
  }

  public render(): React.ReactElement<IToolkitViewerProps> {
    return (
      <div className={styles.toolkitViewer}>
        {/* <button className={ styles.button } onClick={ this._getDocuments }>
          <span className={ styles.label }>get documents</span>
        </button> */}

        <DetailsListDocuments data={this.state.data} />
      </div>
    );
  }
}

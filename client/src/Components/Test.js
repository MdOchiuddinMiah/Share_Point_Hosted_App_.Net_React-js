import * as React from 'react';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Toggle } from 'office-ui-fabric-react/lib/Toggle';
import {
  DetailsList,
  DetailsListLayoutMode,
  Selection,
  IColumn
} from 'office-ui-fabric-react/lib/DetailsList';
import { MarqueeSelection } from 'office-ui-fabric-react/lib/MarqueeSelection';


let _items= [];

const fileIcons = [
  { 'name': 'accdb' },
  { 'name': 'csv' },
  { 'name': 'docx' },
  { 'name': 'dotx' },
  { 'name': 'mpp' },
  { 'name': 'mpt' },
  { 'name': 'odp' },
  { 'name': 'ods' },
  { 'name': 'odt' },
  { 'name': 'one' },
  { 'name': 'onepkg' },
  { 'name': 'onetoc' },
  { 'name': 'potx' },
  { 'name': 'ppsx' },
  { 'name': 'pptx' },
  { 'name': 'pub' },
  { 'name': 'vsdx' },
  { 'name': 'vssx' },
  { 'name': 'vstx' },
  { 'name': 'xls' },
  { 'name': 'xlsx' },
  { 'name': 'xltx' },
  { 'name': 'xsn' }
];

export interface IDetailsListDocumentsExampleState {
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
}

export class Test extends React.Component<any, IDetailsListDocumentsExampleState> {
_selection;

  constructor(props) {
    super(props);

    //  Populate with items for demos.
    if (_items.length === 0) {
      for (let i = 0; i < 500; i++) {
        const randomDate = this._randomDate(new Date(2012, 0, 1), new Date());
        const randomFileSize = this._randomFileSize();
        const randomFileType = this._randomFileIcon();
        let fileName = 'abc'.replace(/\W/g, '');
        let userName = 'bcd'.replace(/[^a-zA-Z ]/g, '');
        fileName = fileName.charAt(0).toUpperCase() + fileName.slice(1).concat(`.${randomFileType.docType}`);
        userName = userName.split(' ').map((name) => name.charAt(0).toUpperCase() + name.slice(1)).join(' ');
        _items.push({
          name: fileName,
          value: fileName,
          iconName: randomFileType.url,
          modifiedBy: userName,
          dateModified: randomDate.dateFormatted,
          dateModifiedValue: randomDate.value,
          fileSize: randomFileSize.value,
          fileSizeRaw: randomFileSize.rawSize
        });
      }
      _items = this._sortItems(_items, 'name');
    }

    const _columns = [
      {
        key: 'column1',
        name: 'File Type',
        headerClassName: 'DetailsListExample-header--FileIcon',
        className: 'DetailsListExample-cell--FileIcon',
        iconClassName: 'DetailsListExample-Header-FileTypeIcon',
        iconName: 'Page',
        isIconOnly: true,
        fieldName: 'name',
        minWidth: 16,
        maxWidth: 16,
        onRender: (item) => {
          return (
            <img
              src={ item.iconName }
              className={ 'DetailsListExample-documentIconImage' }
            />
          );
        }
      },
      {
        key: 'column2',
        name: 'Name',
        fieldName: 'name',
        minWidth: 210,
        maxWidth: 350,
        isRowHeader: true,
        isResizable: true,
        isSorted: true,
        isSortedDescending: false,
        onColumnClick: this._onColumnClick,
        data: 'string',
        isPadded: true
      },
      {
        key: 'column3',
        name: 'Date Modified',
        fieldName: 'dateModifiedValue',
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        onColumnClick: this._onColumnClick,
        data: 'number',
        onRender: (item) => {
          return (
            <span>
              { item.dateModified }
            </span>
          );
        },
        isPadded: true
      },
      {
        key: 'column4',
        name: 'Modified By',
        fieldName: 'modifiedBy',
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        isCollapsable: true,
        data: 'string',
        onColumnClick: this._onColumnClick,
        onRender: (item) => {
          return (
            <span>
              { item.modifiedBy }
            </span>
          );
        },
        isPadded: true
      },
      {
        key: 'column5',
        name: 'File Size',
        fieldName: 'fileSizeRaw',
        minWidth: 70,
        maxWidth: 90,
        isResizable: true,
        isCollapsable: true,
        data: 'number',
        onColumnClick: this._onColumnClick,
        onRender: (item) => {
          return (
            <span>
              { item.fileSize }
            </span>
          );
        }
      },
    ];

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
      isCompactMode: false
    };
  }

  render() {
    let { columns, isCompactMode, items, selectionDetails } = this.state;

    return (
      <div>
      
        <div>{ selectionDetails }</div>
        <TextField
          label='Filter by name:'
          onChanged={ this._onChangeText }
        />
        <MarqueeSelection selection={ this._selection }>
          <DetailsList
            items={ items }
            compact={ isCompactMode }
            columns={ columns }
            setKey='set'
            layoutMode={ DetailsListLayoutMode.justified }
            isHeaderVisible={ true }
            selection={ this._selection }
            selectionPreservedOnEmptyClick={ true }
            onItemInvoked={ this._onItemInvoked }
            enterModalSelectionOnTouch={ true }
          />
        </MarqueeSelection>
      </div>
    );
  }

 componentDidUpdate(previousProps, previousState: IDetailsListDocumentsExampleState) {
    if (previousState.isModalSelection !== this.state.isModalSelection) {
      this._selection.setModal(this.state.isModalSelection);
    }
  }

 _onChangeCompactMode(checked) {
    this.setState({ isCompactMode: checked });
  }

_onChangeModalSelection(checked){
    this.setState({ isModalSelection: checked });
  }

  _onChangeText(text){
    this.setState({ items: text ? _items.filter(i => i.name.toLowerCase().indexOf(text) > -1) : _items });
  }

_onItemInvoked(item) {
    alert(`Item invoked: ${item.name}`);
  }

 _randomDate(start, end){
    const date = new Date(start.getTime() + Math.random() * (end.getTime() - start.getTime()));
    const dateData = {
      value: date.valueOf(),
      dateFormatted: date.toLocaleDateString()
    };
    return dateData;
  }

 _randomFileIcon(){
    const docType = fileIcons[Math.floor(Math.random() * fileIcons.length) + 0].name;
    return {
      docType,
      url: `https://static2.sharepointonline.com/files/fabric/assets/brand-icons/document/svg/${docType}_16x1.svg`
    };
  }

_randomFileSize() {
    const fileSize= Math.floor(Math.random() * 100) + 30;
    return {
      value: `${fileSize} KB`,
      rawSize: fileSize
    };
  }

_getSelectionDetails() {
    let selectionCount = this._selection.getSelectedCount();

    switch (selectionCount) {
      case 0:
        return 'No items selected';
      case 1:
        return '1 item selected: ' + (this._selection.getSelection()[0]).name;
      default:
        return `${selectionCount} items selected`;
    }
  }

_onColumnClick(ev: React.MouseEvent<HTMLElement>, column: IColumn) {
    const { columns, items } = this.state;
    let newItems = items.slice();
    let newColumns= columns.slice();
    let currColumn= newColumns.filter((currCol, idx ) => {
      return column.key === currCol.key;
    })[0];
    newColumns.forEach((newCol) => {
      if (newCol === currColumn) {
        currColumn.isSortedDescending = !currColumn.isSortedDescending;
        currColumn.isSorted = true;
      } else {
        newCol.isSorted = false;
        newCol.isSortedDescending = true;
      }
    });
    newItems = this._sortItems(newItems, currColumn.fieldName, currColumn.isSortedDescending);
    this.setState({
      columns: newColumns,
      items: newItems
    });
  }

_sortItems(items, sortBy, descending = false){
    if (descending) {
      return items.sort((a, b) => {
        if (a[sortBy] < b[sortBy]) {
          return 1;
        }
        if (a[sortBy] > b[sortBy]) {
          return -1;
        }
        return 0;
      });
    } else {
      return items.sort((a, b) => {
        if (a[sortBy] < b[sortBy]) {
          return -1;
        }
        if (a[sortBy] > b[sortBy]) {
          return 1;
        }
        return 0;
      });
    }
  }
}
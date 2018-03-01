import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import {
  Label,
  PrimaryButton,
  TextField,
  Button,
  DialogFooter,
  DialogContent,
  List,
  Check
} from 'office-ui-fabric-react';

import {
  DetailsList,
  DetailsListLayoutMode,
  Selection,
  IColumn
} from 'office-ui-fabric-react/lib/DetailsList';
import { MarqueeSelection } from 'office-ui-fabric-react/lib/MarqueeSelection';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';

import { Dialog } from '@microsoft/sp-dialog';

const _items: any[] = [];

const _columns: IColumn[] = [
  {
    key: 'Gabarit',
    name: 'Name',
    fieldName: 'name',
    minWidth: 100,
    maxWidth: 200,
    isResizable: true,
    ariaLabel: 'Operations for name'
  },
  {
    key: 'column2',
    name: 'Value',
    fieldName: 'value',
    minWidth: 100,
    maxWidth: 200,
    isResizable: true,
    ariaLabel: 'Operations for value'
  },
];


interface IGabaritContentProps {
  message: string;
  newFileName: string;
  close: () => void;
  submit: (filename: string, gabarit: string) => void;
  defaultGabarit?: string;
}

class GabaritPickerDialogContent extends React.Component<IGabaritContentProps, {items: {}[];selectionDetails: {};}>{
  private _gabaritName: string;
  public newFileName: string;
  private _selection: Selection;


  

  

  constructor(props){
    super(props);
    //this.newFileName = "yyyyyy";    

    if (_items.length === 0) {
      for (let i = 0; i < 200; i++) {
        _items.push({
          key: i,
          name: 'Item ' + i,
          value: i
        });
      }
    }

    this._selection = new Selection({
      onSelectionChanged: () => this.setState({ selectionDetails: this._getSelectionDetails() })
    });

    this.state = {
      items: _items,
      selectionDetails: this._getSelectionDetails()
    };



  }
  



  private HandleFileNameChange = (event) => {   
    this.newFileName = event.target.value;      
  }






  public render(): JSX.Element {
    const { items } = this.state;

    return <DialogContent
      title='Select Gabarit'
      subText={this.props.message}
      onDismiss={this.props.close}
      showCloseButton={true}
    >



      <input        
        value={this.newFileName}
        onChange={this.HandleFileNameChange}
        type="text"        
      />
         <TextField
          label='Filter by name:'
         
        />

        <MarqueeSelection selection={this._selection}>
          <DetailsList
            items={ items }
            columns={ _columns }
            setKey='set'
            layoutMode={ DetailsListLayoutMode.fixedColumns }
            selection={ this._selection }
            selectionPreservedOnEmptyClick={ true }
            ariaLabelForSelectionColumn='Toggle selection'
            ariaLabelForSelectAllCheckbox='Toggle selection for all items'
            onItemInvoked={ this._onItemInvoked }
          />
        </MarqueeSelection>
      
      
      <DialogFooter>
        <Button text='Cancel' title='Cancel' onClick={this.props.close} />
        <PrimaryButton text='OK' title='OK' onClick={() => { this.props.submit(this.newFileName,this.newFileName); }} />
      </DialogFooter>
    </DialogContent>;
  }


  private _getSelectionDetails(): string {
    const selectionCount = this._selection.getSelectedCount();

    switch (selectionCount) {
      case 0:
        return 'No items selected';
      case 1:
        return '1 item selected: ' + (this._selection.getSelection()[0] as any).name;
      default:
        return `${selectionCount} items selected`;
    }
  }

  private _onItemInvoked(item: any): void {
    alert(`Item invoked: ${item.name}`);
  }

  //@autobind
  //private _onColorChange(color: string): void {
  //  this._pickedColor = color;
  //}
}


export default class SelectGabaritDialog extends BaseDialog {
  public message: string;
  public gabaritName: string;
  public fileName: string;

  public render(): void {
    ReactDOM.render(<GabaritPickerDialogContent
      close={this.close}
      message={this.message}
      newFileName={this.gabaritName}
      submit={this._submit}
    />, this.domElement);
  }

  public getConfig(): IDialogConfiguration {
    return {
      isBlocking: false
    };
  }

  @autobind
  private _submit(_fileName: string, _gabName: string): void {
    this.gabaritName = _gabName;
    this.fileName = _fileName;
    this.close();
  }
}
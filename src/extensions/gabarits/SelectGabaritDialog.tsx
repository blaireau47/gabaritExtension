import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import { 
  SPHttpClient, ISPHttpClientOptions, 
  SPHttpClientResponse, IHttpClientOptions,
  HttpClientResponse, 
  HttpClient } from 
  '@microsoft/sp-http';
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
  SelectionMode,
  IColumn
} from 'office-ui-fabric-react/lib/DetailsList';
import { MarqueeSelection } from 'office-ui-fabric-react/lib/MarqueeSelection';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';

import { Dialog } from '@microsoft/sp-dialog';

import pnp from "sp-pnp-js";

import { Web,CamlQuery  } from  "sp-pnp-js";



const _columns: IColumn[] = [
  {
    key: 'name',
    name: 'Name',
    fieldName: 'name',
    minWidth: 100,
    maxWidth: 200,
    isResizable: true,
    ariaLabel: 'Operations for name'
  },
  {
    key: 'value',
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
  urlTemplateLibrary: string;  
  gabarits:any[];
  close: () => void;
  submit: (filename: string, gabarit: string) => void;
}

class GabaritPickerDialogContent extends React.Component<IGabaritContentProps, {gabarits: {}[];selectionDetails: {};}>{
  public selectedGabaritName: string;
  public newFileName: string;
  private _selection: Selection;
  private urlWebTemplateLibrary: string;
   

  constructor(props){
    
    super(props);

    /*
    if (this.gabarits.length === 0) {
      for (let i = 0; i < 5; i++) {
        this.gabarits.push({
          key: i,
          name: 'Item ' + i,
          value: i
        });
      }
    }*/

    console.log(this.props.gabarits);
    this._selection = new Selection({
      onSelectionChanged: () => this.setState({ selectionDetails: this._getSelectionDetails() })
    });

    this.state = {
      gabarits: this.props.gabarits,
      selectionDetails: this._getSelectionDetails()
    };

    console.log(this.state.gabarits);      
    console.log(this.props.gabarits.length);


  }
  
  private HandleFileNameChange = (event) => {   
    this.newFileName = event.target.value;      
  }

  public render(): JSX.Element {
    const { gabarits } = this.state;
    console.log("binding selectect list with gabarits " + gabarits)
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
      <MarqueeSelection selection={this._selection}>
        <DetailsList
          items={ gabarits }          
          columns={ _columns }
          setKey='set'
          layoutMode={ DetailsListLayoutMode.fixedColumns }
          selection={ this._selection }
          selectionPreservedOnEmptyClick={ true }
          ariaLabelForSelectionColumn='Toggle selection'
          ariaLabelForSelectAllCheckbox='Toggle selection for all items'
          onItemInvoked={ this._onItemInvoked }
          selectionMode= { SelectionMode.single }
        />
      </MarqueeSelection>            
      <DialogFooter>
        <Button text='Cancel' title='Cancel' onClick={this.props.close} />
        <PrimaryButton text='OK' title='OK' onClick={() => { this.props.submit(this.newFileName,this.selectedGabaritName); }} />
      </DialogFooter>
    </DialogContent>;
  }


  private _getSelectionDetails(): string {
    const selectionCount = this._selection.getSelectedCount();

    switch (selectionCount) {
      case 0:
        this.selectedGabaritName = undefined;
        return 'No items selected';
      case 1:
        this.selectedGabaritName = (this._selection.getSelection()[0] as any).value;
        return (this._selection.getSelection()[0] as any).value;
      default:
        this.selectedGabaritName = undefined;
        return `${selectionCount} items selected`;
    }
  }

  private _onItemInvoked(item: any): void {
    console.log(`Item invoked: ${item.name}`);
  }


}


export default class SelectGabaritDialog extends BaseDialog {
  public message: string;
  public gabaritName: string;
  public fileName: string;
  public urlTemplateLibrary: string;
  public gabarits: any[];

  public render(): void {
    ReactDOM.render(<GabaritPickerDialogContent
      close={this.close}     
      urlTemplateLibrary = {this.urlTemplateLibrary}
      message={this.message}      
      submit={this._submit}      
      gabarits={this.gabarits}
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
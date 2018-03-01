import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import {
  autobind,
  Label,
  PrimaryButton,
  TextField,
  Button,
  DialogFooter,
  DialogContent
} from 'office-ui-fabric-react';
import { Dialog } from '@microsoft/sp-dialog';


interface IGabaritContentProps {
  message: string;
  newFileName: string;
  close: () => void;
  submit: (gabarit: string) => void;
  defaultGabarit?: string;
}

class GabaritPickerDialogContent extends React.Component<IGabaritContentProps, {}>{
  private _gabaritName: string;
  public newFileName: string;
  constructor(props){
    super(props);
    //this.newFileName = "yyyyyy";    
  }

  private HandleFileNameChange = (event) => {   
    this.newFileName = event.target.value;      
  }

  //HandleFileNameChange(event) {
  //  Dialog.alert("Change bitcha");
  //  this.newFileName = event.target.value;
  //};

  public render(): JSX.Element {
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
      
      
      <DialogFooter>
        <Button text='Cancel' title='Cancel' onClick={this.props.close} />
        <PrimaryButton text='OK' title='OK' onClick={() => { this.props.submit(this.newFileName); }} />
      </DialogFooter>
    </DialogContent>;
  }

  //@autobind
  //private _onColorChange(color: string): void {
  //  this._pickedColor = color;
  //}
}


export default class SelectGabaritDialog extends BaseDialog {
  public message: string;
  public gabaritName: string;

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
  private _submit(gabName: string): void {
    this.gabaritName = gabName;
    this.close();
  }
}
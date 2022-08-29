import { ServiceScope } from '@microsoft/sp-core-library';
import { BaseDialog, IDialogConfiguration } from '@microsoft/sp-dialog';
import { DefaultButton } from 'office-ui-fabric-react';
import * as React from 'react';
import * as ReactDOM from 'react-dom';
import HelloWorld from './HelloWorld';

export default class HelloWorldDialog extends BaseDialog {
    private _serviceScope: ServiceScope;
    private _siteUrl: string;
       
    public constructor(serviceScope: ServiceScope, siteUrl: string) {
      super({ isBlocking: true });
      
      this._serviceScope = serviceScope;
      this._siteUrl = siteUrl;
    }
  
    public render(): void {
      
      const element = <div style={{ width: 1000, padding: 20 }}>
        <HelloWorld serviceScope={this._serviceScope} siteUrl={this._siteUrl} />
        <DefaultButton text='Close' onClick={this.close}/>
      </div>;

      
      ReactDOM.render(element, this.domElement);
    }
    
    public getConfig(): IDialogConfiguration {
      return { isBlocking: false };
    }
  
    protected onAfterClose(): void {
      super.onAfterClose();
  
      // Clean up the element for the next dialog
      ReactDOM.unmountComponentAtNode(this.domElement);
    }  
  }
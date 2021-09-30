import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './JQueryWebpartWebPart.module.scss';
import * as strings from 'JQueryWebpartWebPartStrings';
import * as $ from 'jquery';
import 'jqueryui';
import {SPComponentLoader} from '@microsoft/sp-loader';

export interface IJQueryWebpartWebPartProps {
  description: string;
}

export default class JQueryWebpartWebPart extends BaseClientSideWebPart<IJQueryWebpartWebPartProps> {
constructor(){
  super();
  SPComponentLoader.loadCss('code.jquery.com/ui/1.11.4/themes/smoothness/jquery-ui.css')
}

  public render(): void {
  
    this.domElement.innerHTML = `
    <div class="accordian">    
    <h3>Step 1</h3>
    <div>
    <p> Para 1</p>
    </div>    
    <h3>Step 2</h3>
    <div>
    <p> Para 2</p>
    </div>   
    <h3>Step 3</h3>
    <div>
    <p>Para 3</p>
    </div>
    </div>` 
    ;
    $(".accordian").accordion();
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}

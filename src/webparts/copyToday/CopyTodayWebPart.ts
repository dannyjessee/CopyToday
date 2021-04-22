import { Version, Environment, EnvironmentType } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './CopyTodayWebPart.module.scss';
import * as strings from 'CopyTodayWebPartStrings';

import { SPComponentLoader } from '@microsoft/sp-loader';

export interface ICopyTodayWebPartProps {
  description: string;
}

export default class CopyTodayWebPart extends BaseClientSideWebPart<ICopyTodayWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = '<div id="siteContent"></div>';

    let htmlContext = this;

    if ((Environment.type != EnvironmentType.ClassicSharePoint) && (typeof SP == 'undefined')) {
      console.log(Environment.type);
      this.loadSPDependencies().then(e => {
        ExecuteOrDelayUntilScriptLoaded(() => { htmlContext.copyToday(htmlContext); }, "sp.js");
      });
    } else {
      console.log(Environment.type);
      ExecuteOrDelayUntilScriptLoaded(() => { htmlContext.copyToday(htmlContext); }, "sp.js");
    }  
    
  }

  public copyToday(htmlContext): void {
    // Get the Title and Miles from the existing item, copy it to a new item in the Running list with today's date
    const existingItemId: number = parseInt(this.getQueryVariable('ItemID'));

    const ctx: SP.ClientContext = new SP.ClientContext(this.context.pageContext.web.absoluteUrl);  
    const list: SP.List = ctx.get_web().get_lists().getByTitle('Running');
    const listItem: SP.ListItem = list.getItemById(existingItemId);
    ctx.load(listItem);  
    ctx.executeQueryAsync((sender, args) => {
      const itemCreateInfo: SP.ListItemCreationInformation = new SP.ListItemCreationInformation();
      const newListItem: SP.ListItem = list.addItem(itemCreateInfo);            
      newListItem.set_item('Title', listItem.get_item('Title'));
      newListItem.set_item('Date', new Date().toUTCString());  // Today
      newListItem.set_item('Miles', listItem.get_item('Miles'));
      newListItem.update();
      
      ctx.load(newListItem);            
      ctx.executeQueryAsync(() => {
        let content:string = '<div>Item successfully copied to today!</div>';
        htmlContext.domElement.querySelector("#siteContent").innerHTML = content; 
      }, (a, b) => {
          alert('Request failed. ' + b.get_message() + '\n' + b.get_stackTrace());
      });     
    },  
    (c, d) => {  
      console.log(d.get_message());  
    });  
    
  }

  public loadSPDependencies(): Promise<{}> {
    return SPComponentLoader.loadScript('/_layouts/15/init.js', {
      globalExportsName: '$_global_init'
    })
      .then((): Promise<{}> => {
        return SPComponentLoader.loadScript('/_layouts/15/MicrosoftAjax.js', {
          globalExportsName: 'Sys'
        });
      })
      .then((): Promise<{}> => {
        return SPComponentLoader.loadScript('/_layouts/15/SP.Runtime.js', {
          globalExportsName: 'SP'
        });
      })
      .then((): Promise<{}> => {
        return SPComponentLoader.loadScript('/_layouts/15/SP.js', {
          globalExportsName: 'SP'
        });
      });

  }
  
  public getQueryVariable(variable) : string {
    var query = window.location.search.substring(1);
    var vars = query.split('&');
    for (var i = 0; i < vars.length; i++) {
        var pair = vars[i].split('=');
        if (decodeURIComponent(pair[0]) == variable) {
            return decodeURIComponent(pair[1]);
        }
    }
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

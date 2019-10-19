import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import styles from './ItRequestsWebPart.module.scss';
import * as strings from 'ItRequestsWebPartStrings';
import 'jquery';
import 'datatables.net';
import 'moment';
import './moment-plugin';
var $: any = (window as any).$;

export interface IItRequestsWebPartProps {
  description: string;
  listName: string;
}

export default class ItRequestsWebPart extends BaseClientSideWebPart<IItRequestsWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
    <link rel="stylesheet" type="text/css" href="https://cdn.datatables.net/1.10.15/css/jquery.dataTables.min.css" />
    <table id="requests" class="display ${styles.itRequests}" cellspacing="0" width="100%">
        <thead>
            <tr>
                <th>ID</th>
                <th>Business unit</th>
                <th>Category</th>
                <th>Status</th>
                <th>Due date</th>
                <th>Assigned to</th>
            </tr>
        </thead>
    </table>`;

    $(document).ready(() => {
      var listnm=this.properties.listName; //Use this variable in REST API when you need to fetch list name from property pane
      $('#requests').DataTable({
          'ajax': {
              'url': "../../_api/web/lists/getbytitle('IT Requests')/items?$select=ID,BusinessUnit,Category,Status,DueDate,AssignedTo/Title&$expand=AssignedTo/Title",
              'headers': { 'Accept': 'application/json;odata=nometadata' },
              'dataSrc': function (data) {
                  return data.value.map(function (item) {
                      return [
                          item.ID,
                          item.BusinessUnit,
                          item.Category,
                          item.Status,
                          new Date(item.DueDate),
                          item.AssignedTo.Title
                      ];
                  });
              }
          },
          columnDefs: [{
              targets: 4,
              render: $.fn.dataTable.render.moment('YYYY/MM/DD')
          }]
      });
  });
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
                PropertyPaneTextField('listName', {
                  label: strings.ListNameFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
    protected get disableReactivePropertyChanges(): boolean {
      return true;
  }
}

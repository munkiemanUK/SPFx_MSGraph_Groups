import * as React from 'react';
import styles from './PnPGraph.module.scss';
import { IPnPGraphProps } from './IPnPGraphProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient, ISPHttpClientOptions } from '@microsoft/sp-http';

import {sp} from '@pnp/sp/';
import { Group } from '@microsoft/microsoft-graph-types';
import { graph } from '@pnp/graph/presets/all';
import "@pnp/graph/groups";
import "@pnp/graph/members";
import * as $ from 'jquery';
require('bootstrap');
require('../../../../node_modules/bootstrap/dist/css/bootstrap.css');
import {
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  IDetailsList
} from 'office-ui-fabric-react/lib/DetailsList';
export interface IState {
  groups: Group[];
  members: any[];
  isAdmin: boolean;
  siteurl: string;
  userId: string;
}

const _columns: IColumn[] = [
  {
    key: 'id',
    name: 'Id',
    fieldName: 'id',
    minWidth: 50,
    maxWidth: 100,
    isResizable: true
  },
  {
    key: 'name',
    name: 'Name',
    fieldName: 'displayName',
    minWidth: 50,
    maxWidth: 150,
    isResizable: true
  },
  {
    key: 'created',
    name: 'Created',
    fieldName: 'createdDateTime',
    minWidth: 50,
    maxWidth: 200,
    isResizable: true
  }
];

export default class PnPGraph extends React.Component<IPnPGraphProps, IState, IColumn> {

  constructor(props: any) {
    super(props);

    this.state = {
      groups: null,
      members: null,
      isAdmin: false,
      siteurl: "",
      userId: ""
    };
  }

  
  public async componentDidMount(): Promise<void> {
    //https://maximusunitedkingdom.sharepoint.com/_api/web/currentuser 
    let tempID="";

    graph.groups.get<Group[]>().then(groups => {
      this.setState({
        groups
      });
    });

    $.ajax({ 
      url: 'https://maximusunitedkingdom.sharepoint.com/_api/web/currentuser', 
      type: "GET", 
      headers:{'Accept': 'application/json; odata=verbose;'}, 
      success: function(data) {
        console.log(data.d.results); 
        //tempID=data.d.results; 
      }, 
      error : function(jqXHR, textStatus, errorThrown) { 
      } 
    }); 

    this._getMembers("da85cb9b-8ae9-4ee9-aa51-a32d61bb08e2");
    let adminAccess = false;

    //const response2 = await this.props.myhttp.get(this.props.mysite+"/_api/web/CurrentUser", SPHttpClient.configurations.v1);
    //if (!(response2.ok)) {throw new Error(await response2.text()); }
    //const responseJSONuser: any = await response2.json();

    //const response3 = await this.props.myhttp.get(this.props.mysite+"/_api/web/GetUserById("+this.state.userId+")/Groups", SPHttpClient.configurations.v1);
    //if (!(response3.ok)) {throw new Error(await response3.text()); }
    //const responseJSONgroupall: any = await response3.json();

    //let responseJSONrecord = responseJSONgroupall.value.filter((item: { Title: string | string[]; }) => item.Title.indexOf("Finance PO App Admins")>-1);
    //if(responseJSONrecord.length>0){
    //  adminAccess=true;
    //}
    //this.setState({isAdmin:adminAccess});
  }

  private async _getMembers(id: string) {   
    const memberList = await graph.groups.getById(id).members();
    //this.setState({members:memberList});
    console.log("group members="+memberList);

    return memberList;
  }

  public render(): React.ReactElement<IPnPGraphProps> {
    
    if (!this.state.groups) {
      return <div>Loading...</div>;
    }

    return (
      <div className="container">
        <h2>Groups at your tenant:</h2>
        <DetailsList
          items={this.state.groups}
          columns={_columns}
          setKey="set"
          layoutMode={DetailsListLayoutMode.justified}
          selectionPreservedOnEmptyClick={true}
          ariaLabelForSelectionColumn="Toggle selection"
          ariaLabelForSelectAllCheckbox="Toggle selection for all items"
        />
        <div>Is Admin {this.state.isAdmin.valueOf}</div>
      </div>
    );
  }
}

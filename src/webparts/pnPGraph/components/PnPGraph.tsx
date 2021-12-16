import * as React from 'react';
import styles from './PnPGraph.module.scss';
import { IPnPGraphProps } from './IPnPGraphProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { Group } from '@microsoft/microsoft-graph-types';
import { graph } from '@pnp/graph';
import "@pnp/graph/groups";
import "@pnp/graph/members";
import {
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  IDetailsList
} from 'office-ui-fabric-react/lib/DetailsList';

export interface IState {
  groups: Group[];
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
      groups: null
    };
  }

  public checkGroupID = (column: IColumn) => { 
    
  }

  public componentDidMount(): void {
    graph.groups.get<Group[]>().then(groups => {
      this.setState({
        groups
      });
    });
  }

  public render(): React.ReactElement<IPnPGraphProps> {
    if (!this.state.groups) {
      return <div>Loading...</div>;
    }

    return (
      <div>
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
      </div>
    );
  }
}

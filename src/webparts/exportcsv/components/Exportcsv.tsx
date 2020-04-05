import * as React from 'react';
import styles from './Exportcsv.module.scss';
import { IExportcsvProps } from './IExportcsvProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { CommandBarButton } from 'office-ui-fabric-react/lib/Button';
import { DetailsList, DetailsListLayoutMode, IColumn, SelectionMode } from 'office-ui-fabric-react/lib/DetailsList';
import { CSVLink } from "react-csv";
import { SPHttpClient } from '@microsoft/sp-http';

export interface IUserRequest {
  displayName: string;
  mail: string;
}

export interface IExportcsvState {
  userRequests: IUserRequest[];
}

export default class Exportcsv extends React.Component<IExportcsvProps, IExportcsvState> {
  private userRequests: IUserRequest[] = [];
  constructor(props: IExportcsvProps) {
    super(props);    
    this.state = {
      userRequests: this.userRequests
    };
  }
  public componentDidMount() {
    const restApi = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/GetByTitle('Requests')/items?$select=Title,Requestor,RequestDate,Location,Status`;
    this.props.context.spHttpClient.get(restApi, SPHttpClient.configurations.v1)
      .then(resp => { return resp.json(); })
      .then(items => {
        this.setState({userRequests: items.value});
      });
  }

  public render(): React.ReactElement<IExportcsvProps> {
    const columns: IColumn[] = [
      {
        key: 'column1',
        name: 'Title',
        fieldName: 'Title',
        minWidth: 100,
        maxWidth: 400,
        isResizable: true
      },
      {
        key: 'column2',
        name: 'Requestor',
        fieldName: 'Requestor',
        minWidth: 100,
        maxWidth: 400,
        isResizable: true
      },
      {
        key: 'column3',
        name: 'Location',
        fieldName: 'Location',
        minWidth: 100,
        maxWidth: 400,
        isResizable: true
      },
      {
        key: 'column4',
        name: 'RequestDate',
        fieldName: 'RequestDate',
        minWidth: 100,
        maxWidth: 400,
        isResizable: true
      },
      {
        key: 'column5',
        name: 'Status',
        fieldName: 'Status',
        minWidth: 100,
        maxWidth: 400,
        isResizable: true
      },
    ];
    return (
      <div className={ styles.exportcsv }>
        <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <CSVLink data={this.state.userRequests} filename={'Requests.csv'}>
                  <CommandBarButton  iconProps={{ iconName: 'ExcelLogoInverse' }} text='Export to Excel' />
              </CSVLink>
            </div>
          </div><br/>
          <DetailsList
              items={this.state.userRequests}
              columns={columns}
              isHeaderVisible={true}
              setKey='set'
              layoutMode={DetailsListLayoutMode.justified}
              selectionMode={SelectionMode.none}
            />
        </div>
      </div>
    );
  }
}

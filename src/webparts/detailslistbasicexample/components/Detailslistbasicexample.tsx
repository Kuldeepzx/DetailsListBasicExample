import * as React from 'react';
import styles from './Detailslistbasicexample.module.scss';
import { Announced } from 'office-ui-fabric-react/lib/Announced';
import { TextField, ITextFieldStyles } from 'office-ui-fabric-react/lib/TextField';
import { DetailsList, DetailsListLayoutMode, Selection, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { MarqueeSelection } from 'office-ui-fabric-react/lib/MarqueeSelection';
import { Fabric } from 'office-ui-fabric-react/lib/Fabric';
import { mergeStyles } from 'office-ui-fabric-react/lib/Styling';

import { IDetailslistbasicexampleProps } from './IDetailslistbasicexampleProps';
import { escape } from '@microsoft/sp-lodash-subset';



// Sp Pnp Setup
import { sp } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { IItemAddResult } from "@pnp/sp/items";
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';

const exampleChildClass = mergeStyles({
  display: 'block',
  marginBottom: '10px',
});

const textFieldStyles: Partial<ITextFieldStyles> = { root: { maxWidth: '300px' } };

export interface IDetailsListBasicExampleItem {
  // Step:-4 ahiya Title String pass thay che.
  Title:string
}


export interface IDetailsListBasicExampleState {
  // Step:-3 ahiya to IDetailsListBaicExampleItem che Step:-4
  items:any;
  selectionDetails: string;

}
export interface IListItem {
  title:string;
}


export default class Detailslistbasicexample extends React.Component<IDetailslistbasicexampleProps, IDetailsListBasicExampleState> {
  private _selection: Selection;
  private _allItems: IDetailsListBasicExampleItem[];
  private _columns: IColumn[];

  constructor(props: IDetailslistbasicexampleProps) {
    super(props);

    this._selection = new Selection({
      onSelectionChanged: () => this.setState({ selectionDetails: this._getSelectionDetails() }),
    });

    // Populate with items for demos.
    this._allItems = [];
    // for (let i = 0; i < this._allItems.length; i++) {
    //   this._allItems.push({
    //     key: i,
    //     name: 'Item ' + i,
    //     value: i,
    //   });
    // }

    console.log(this._allItems);

    this._columns = [
      { key: 'ID', name: 'ID', fieldName: 'ID', minWidth: 100, maxWidth: 200, isResizable: true },
       { key: 'Title', name: 'Value', fieldName: 'Title', minWidth: 100, maxWidth: 200, isResizable: true },
       { key: 'Region', name: 'Region', fieldName: 'region', minWidth: 100, maxWidth: 200, isResizable: true },

    ];

    // console.log(this._columns)

    this.state = {
      items:[],
      selectionDetails: this._getSelectionDetails(),
    
    };

    if (Environment.type === EnvironmentType.SharePoint) {
      this._getListItems();
    }
    else if (Environment.type === EnvironmentType.Local) {
      // return (<div>Whoops! you are using local host...</div>);
    }
  }

  public render(): React.ReactElement<IDetailslistbasicexampleProps> {
    // step:-2 ahiya items state ma che. Go to step:-3
    const { items, selectionDetails } = this.state;
    
    return (
      <Fabric>
      <div className={exampleChildClass}>{selectionDetails}</div>
      <Announced message={selectionDetails} />
      <TextField
        className={exampleChildClass}
        label="Filter by name:"
        onChange={this._onFilter}
        styles={textFieldStyles}
      />
       {/* <button onClick = {this._getItems} > Click Me </button> */}
      <Announced message={`Number of items after filter applied: ${items.length}.`} />
      <MarqueeSelection selection={this._selection}>
        <DetailsList
            //step 1:-   Items props che Go to step:-2
          items={items}
          columns={this._columns}
          setKey="set"
          layoutMode={DetailsListLayoutMode.justified}
          selection={this._selection}
          selectionPreservedOnEmptyClick={true}
          ariaLabelForSelectionColumn="Toggle selection"
          ariaLabelForSelectAllCheckbox="Toggle selection for all items"
          checkButtonAriaLabel="Row checkbox"
          onItemInvoked={this._onItemInvoked}
        />
      </MarqueeSelection>
     
    </Fabric>
    );
  }
  
  async _getListItems() {
    const allDetails: any[] = await sp.web.lists.getByTitle("vendor").items.getAll();
    console.log(allDetails);
    
    this.setState({items:allDetails});
   
    
  }
// async _getItems(){
//   const allItems : any[] = await sp.web.lists.getByTitle("vendor").items.getAll();
//   console.log(allItems);
// }


  private  _getSelectionDetails(): any {
    const selectionCount = this._selection.getSelectedCount();
    const allITems = sp.web.lists.getByTitle("vendor").items.getAll();
    console.log(allITems);

    console.log(selectionCount);
    // switch (selectionCount) {
    //   case 0:
    //     return 'No items selected';
    //   case 1:
    //     return '1 item selected: ' + (this._selection.getSelection()[0] as IDetailsListBasicExampleItem).name;
    //   default:
    //     return `${selectionCount} items selected`;
    // }
  }
  private _onFilter = (ev: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, text: string): void => {
    // this.setState({
    //   items: text ? this._allItems.filter(i => i.name.toLowerCase().indexOf(text) > -1) : this._allItems,
    // });
  };

  private _onItemInvoked = (item: IDetailsListBasicExampleItem): void => {
    // alert(`Item invoked: ${item.name}`);
  };
}

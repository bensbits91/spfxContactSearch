import * as React from 'react';
import { ShimmeredDetailsList } from 'office-ui-fabric-react/lib/ShimmeredDetailsList';
import { buildColumns, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { Link } from 'office-ui-fabric-react/lib/Link';





export interface IResult {
   key: string;
   FirstName: string;
   Title: string;
   JobTitle: string;
   Organization: string;
   Company: string;
   Division: string;
   Program: string;
   Email: string;
   WorkPhone: string;
   WorkAddress: string;
   CellPhone: string;
}


export interface IListProp {
   handler: any;
   items?: IResult[];
   searchTerms?: any;
   // order?: string;
   show_department: boolean;
   show_division: boolean;
   show_organization: boolean;
}

export interface IListState {
   // sortedItems: any;
   columns: IColumn[];
   // searchTerms?: any;
   // order?: string;
   // show_department: boolean;
   // show_division: boolean;
   // show_organization: boolean;
}






export default class List extends React.Component<IListProp, IListState> {

   constructor(props) {
      super(props);

      this.state = {
         // sortedItems: this.props.items,
         columns: _buildColumns(this.props.items, this.props.show_organization, this.props.show_department, this.props.show_division),
         // searchTerms: this.props.searchTerms,
         // order: this.props.order,
         // show_organization: this.props.show_organization,
         // show_department: this.props.show_department,
         // show_division: this.props.show_division
      };

      this._renderItemColumn = this._renderItemColumn.bind(this);
   }

   public render() {
      const { columns } = this.state;
      console.log('%c : List -> render -> this.state', 'color:yellow', this.state);
      const { items, searchTerms } = this.props;
      console.log('%c : List -> render -> this.props', 'color:yellow', this.props);
      columns.map(column => {
         column.isResizable = true;
         column.name = column.fieldName == 'Company' ? 'Department'
         : column.fieldName == 'Title' ? 'Last Name'
         : column.fieldName.replace(/([A-Z])/g, ' $1').trim();
      });
      
      return (
         <ShimmeredDetailsList
            items={items}
            // items={sortedItems}
            setKey="set"
            columns={columns}
            onRenderItemColumn={this._renderItemColumn}
            onColumnHeaderClick={this._onColumnClick}
            ariaLabelForSelectionColumn="Toggle selection"
            ariaLabelForSelectAllCheckbox="Toggle selection for all items"
            checkButtonAriaLabel="Row checkbox"
            searchTerms={searchTerms}
         />
      );
   }

   // public sendData = (order) => {
   //   this.props.handler(order);
   // }

   // public componentDidUpdate(previousProps: IListProp, previousState: IListState) {
   //    if (previousState.sortedItems != this.props.items) {
   //       this.setState({ sortedItems: this.props.items });
   //    }
   //    if (previousState.order != this.props.order) {
   //       this.setState({ order: this.props.order });
   //    }
   //    if (previousState.show_organization != this.props.show_organization) {
   //       this.setState({
   //          show_organization: this.props.show_organization,
   //          columns: _buildColumns(this.props.items, this.props.show_organization, this.props.show_department, this.props.show_division)
   //       });
   //    }
   //    if (previousState.show_department != this.props.show_department) {
   //       this.setState({
   //          show_department: this.props.show_department,
   //          columns: _buildColumns(this.props.items, this.props.show_organization, this.props.show_department, this.props.show_division)
   //       });
   //    }
   //    if (previousState.show_division != this.props.show_division) {
   //       this.setState({
   //          show_division: this.props.show_division,
   //          columns: _buildColumns(this.props.items, this.props.show_organization, this.props.show_department, this.props.show_division)
   //       });
   //    }
   // }


   public _renderItemColumn(item: IResult, index: number, column: IColumn/* , searchTerms: any */) {
      const searchTermsToHighlight = this.props.searchTerms;

      let highlightHits = (str) => {
         for (let term of searchTermsToHighlight) {
            const searchTermRegex = new RegExp(term.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'), "ig");
            const searchTermHighlighted = '<span style="background-color:yellow;">$&</span>';
            str = str.replace(searchTermRegex, searchTermHighlighted);
         }
         return str;
      };

      const fieldContent = item[column.fieldName as keyof IResult] as string;

      switch (column.key) {
         case 'FirstName': case 'Title':
            return (
               fieldContent != null
                  ? <Link href={"https://delve-gcc.office.com/?p=" + item.Email + "&v=work"} target="about:blank">
                     <span dangerouslySetInnerHTML={{ __html: highlightHits(fieldContent) }} />
                  </Link>
                  : ''
            );
         case 'JobTitle': case 'Organization': case 'Company': case 'Division': case 'Program':
            return (
               fieldContent != null
                  ? <span dangerouslySetInnerHTML={{ __html: highlightHits(fieldContent) }} />
                  : ''
            );
         default:
            return <span>{fieldContent}</span>;
      }
   }

   private _onColumnClick = (event: React.MouseEvent<HTMLElement>, column: IColumn): void => {
      console.log('%c : List -> event', 'color:indigo', event);
      this.props.handler(column.fieldName);

      // let isSortedDescending = column.isSortedDescending;

      // if (column.isSorted) {
      //    isSortedDescending = !isSortedDescending;
      // }

      // this.setState({
      //    order: column.fieldName
      // }, () => {
      // this.props.handler(this.state.order);
      //  this.sendData(this.state.order);
      //    });
   }

}

function _buildColumns(items: IResult[], show_organization, show_department, show_division): IColumn[] {
   let theColumns = [];
   items.map(item => {
      theColumns.push({
         FirstName: item.FirstName,
         Title: item.Title,
         JobTitle: item.JobTitle,
         WorkPhone: item.WorkPhone,
         CellPhone: item.CellPhone,
         Email: item.Email,
         ...show_department ? { Company: item.Company } : null,
         ...show_division ? { Division: item.Division } : null,
         ...show_organization ? { Organization: item.Organization } : null,
         Program: item.Program,
         WorkAddress: item.WorkAddress
      });
   });
   const columns = buildColumns(theColumns);
   return columns;
}
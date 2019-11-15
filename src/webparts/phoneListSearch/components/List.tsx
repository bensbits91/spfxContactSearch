import * as React from 'react';
import { ShimmeredDetailsList } from 'office-ui-fabric-react/lib/ShimmeredDetailsList';
import { buildColumns, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { Link } from 'office-ui-fabric-react/lib/Link';


/* export  */interface IResult {
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

/* export  */interface IListProp {
   handler: any;
   items?: IResult[];
   searchTerms?: any;
   show_department: boolean;
   show_division: boolean;
   show_organization: boolean;
}

/* export  */interface IListState {
   columns: IColumn[];
}

// might be able to make this component stateless by moving _buildColumns to App.tsx  <------------------------------------------------




export default class List extends React.Component<IListProp, IListState> {

   constructor(props) {
      super(props);

      this.state = {
         columns: _buildColumns(this.props.items, this.props.show_organization, this.props.show_department, this.props.show_division),
      };

      this._renderItemColumn = this._renderItemColumn.bind(this);
   }

   public render() {
      const { columns } = this.state;
      const { items, searchTerms } = this.props;
      columns.map(column => {
         column.isResizable = true;
         column.name = column.fieldName == 'Company' ? 'Department'
            : column.fieldName == 'Title' ? 'Last Name'
               : column.fieldName.replace(/([A-Z])/g, ' $1').trim();
      });

      return (
         <ShimmeredDetailsList
            items={items}
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
      this.props.handler(column.fieldName);
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
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneCheckbox, PropertyPaneToggle } from '@microsoft/sp-property-pane';

import * as strings from 'PhoneListSearchWebPartStrings';
import PhoneListSearch from './components/PhoneListSearch';
import { IPhoneListSearchProps } from './components/IPhoneListSearchProps';

import { ITheme, mergeStyleSets, getTheme, getFocusStyle, noWrap } from 'office-ui-fabric-react/lib/Styling';
import { IIconProps } from 'office-ui-fabric-react/lib/Icon';
import { SearchBox } from 'office-ui-fabric-react/lib/SearchBox';
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { CommandBar } from 'office-ui-fabric-react/lib/CommandBar';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { DetailsList, buildColumns, IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { PersonaCoin, PersonaSize } from 'office-ui-fabric-react/lib/Persona';
import { Facepile, IFacepilePersona, IFacepileProps } from 'office-ui-fabric-react/lib/Facepile';
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import styles from './components/PhoneListSearch.module.scss';
import { debounce } from 'lodash';
import { Dropdown, DropdownMenuItemType, IDropdownStyles, IDropdownOption } from 'office-ui-fabric-react/lib/Dropdown';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { ShimmeredDetailsList } from 'office-ui-fabric-react/lib/ShimmeredDetailsList';

import { graph } from "@pnp/graph";




export interface IPhoneListSearchWebPartProps {
  appHeading: string;
  searchBoxPlaceholder: string;
  initialResultText: string;
  noResultText: string;
  showOrganization: boolean;
  showDepartment: boolean;
  showDivision: boolean;
}

export interface IMainAppProps {
  appHeading: string;
  searchBoxPlaceholder: string;
  initialResultText: string;
  noResultText: string;
  showOrganization: boolean;
  showDepartment: boolean;
  showDivision: boolean;
}

export interface IMainAppState {
  needUpdate: boolean;
  items: any;
  searchTerms?: string;
  view?: string;
  order?: string;
  size?: string;
  showPanel: boolean;
  filters: string;
  hasFiltersOrganization: boolean;
  hasFiltersDepartment: boolean;
  hasFiltersDivision: boolean;
  clearFilters: boolean;
}

export interface IContactSearchBoxProps {
  parentCallback;
  view?: string;
  order?: string;
  size?: string;
  showPanel: boolean;
  filters: string;
  hasFiltersOrganization: boolean;
  hasFiltersDepartment: boolean;
  hasFiltersDivision: boolean;
  searchBoxPlaceholder: string;
  showOrganization: boolean;
  showDepartment: boolean;
  showDivision: boolean;
  clearFilters: boolean;
}

export interface IContactSearchBoxState {
  searchTerms: any;
  items: any;
  view?: string;
  order?: string;
  size?: string;
  needUpdate: boolean;
  showPanel: boolean;
  filters: string;
  hasFiltersOrganization: boolean;
  hasFiltersDepartment: boolean;
  hasFiltersDivision: boolean;
  showOrganization: boolean;
  showDepartment: boolean;
  showDivision: boolean;
  clearFilters: boolean;
}


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
}

export interface IContactCardGridProps {
  items?: any;
  searchTerms: string;
  size?: string;
  showOrganization: boolean;
  showDepartment: boolean;
  showDivision: boolean;
}

export interface IContactCardGridState {
  items?: any;
  showOrganization: boolean;
  showDepartment: boolean;
  showDivision: boolean;
}

export interface IContactCardProps {
  item?: any;
  searchTerms: string;
  size?: string;
  showOrganization: boolean;
  showDepartment: boolean;
  showDivision: boolean;
}

export interface IContactCardState {
  item?: any;
  showOrganization: boolean;
  showDepartment: boolean;
  showDivision: boolean;
}

export interface IFacepileBasicExampleState {
  numberOfFaces: any;
  imagesFadeIn: boolean;
  personaSize: PersonaSize;
}

export interface IFacepileBasicExampleProps {
  personas: any;
  personaSize: number;
}

export interface ICommandBarSearchControlsProps {
  parentCallback;
  view?: string;
  order?: string;
  size?: string;
  showPanel: boolean;
  filters: string;
  clearFilters: boolean;
}

export interface ICommandBarSearchControlsState {
  view?: string;
  order?: string;
  size?: string;
  showPanel: boolean;
  filters: string;
  hasFiltersOrganization: boolean;
  hasFiltersDepartment: boolean;
  hasFiltersDivision: boolean;
  clearFilters: boolean;
}

export interface IDetailsListCustomColumnsResultsProp {
  parentCallback;
  items?: IResult[];
  searchTerms?: any;
  order?: string;
  showOrganization: boolean;
  showDepartment: boolean;
  showDivision: boolean;
}

export interface IDetailsListCustomColumnsResultsState {
  sortedItems: any;
  columns: IColumn[];
  searchTerms?: any;
  order?: string;
  showOrganization: boolean;
  showDepartment: boolean;
  showDivision: boolean;
}

export interface IFilterPanelProps {
  parentCallback?;
  showPanel: boolean;
  filters: string;
  clearFilters: boolean;
}

export interface IFilterPanelState {
  showPanel: boolean;
  hasChoiceData: boolean;
  filters: string;
  filtersOrganization: any;
  filtersDepartment: any;
  filtersDivision: any;
  clearFilters: boolean;
}

export interface IDropdownControlledMultiState {
  selectedItems: string[];
}

export interface IDropdownControlledMultiProps {
  choices?: any;
  label: string;
  placeholder: string;
  onChange: any;
}



const theme: ITheme = getTheme();
const { palette, fonts } = theme;
const classNames = mergeStyleSets({
  fileIconHeaderIcon: {
    padding: 0,
    fontSize: '16px'
  },
  fileIconCell: {
    textAlign: 'center',
    selectors: {
      '&:before': {
        content: '.',
        display: 'inline-block',
        verticalAlign: 'middle',
        height: '100%',
        width: '0px',
        visibility: 'hidden'
      }
    }
  },
  fileIconImg: {
    verticalAlign: 'middle',
    maxHeight: '16px',
    maxWidth: '16px'
  },
  controlWrapper: {
    display: 'flex',
    flexWrap: 'wrap'
  },
  exampleToggle: {
    display: 'inline-block',
    marginBottom: '10px',
    marginRight: '30px'
  },
  selectionDetails: {
    marginBottom: '20px'
  },
  listGridExample: {
    overflow: 'hidden',
    fontSize: 0,
    position: 'relative'
  },
  listGridExampleTile: {
    textAlign: 'center',
    outline: 'none',
    position: 'relative',
    float: 'left',
    background: palette.neutralLighter,
    selectors: {
      'focus:after': {
        content: '',
        position: 'absolute',
        left: 2,
        right: 2,
        top: 2,
        bottom: 2,
        boxSizing: 'border-box',
        border: `1px solid ${palette.white}`
      }
    }
  },
  listGridExampleSizer: {
    paddingBottom: '100%'
  },
  listGridExamplePadder: {
    position: 'absolute',
    left: 2,
    top: 2,
    right: 2,
    bottom: 2
  },
  listGridExampleLabel: {
    background: 'rgba(0, 0, 0, 0.3)',
    color: '#FFFFFF',
    position: 'absolute',
    padding: 10,
    bottom: 0,
    left: 0,
    width: '100%',
    fontSize: fonts.small.fontSize,
    boxSizing: 'border-box'
  },
  listGridExampleImage: {
    position: 'absolute',
    top: 0,
    left: 0,
    width: '100%'
  },
  listGridExampleContent: {
    fontSize: 14,
    left: 0,
    position: 'absolute',
    top: 0,
    width: '100%'
  },
  itemCell: [
    getFocusStyle(theme, { inset: -1 }),
    {
      minHeight: 54,
      padding: 10,
      boxSizing: 'border-box',
      borderBottom: `1px solid #aaa`,
      display: 'flex',
      selectors: {
        '&:hover': { background: palette.neutralLight }
      }
    }
  ],
  itemImage: {
    flexShrink: 0
  },
  itemContent: {
    marginLeft: 10,
    overflow: 'hidden',
    flexGrow: 1
  },
  itemName: [
    fonts.xLarge,
    {
      whiteSpace: 'nowrap',
      overflow: 'hidden',
      textOverflow: 'ellipsis'
    }
  ],
  itemIndex: {
    fontSize: fonts.small.fontSize,
    color: palette.neutralTertiary,
    marginBottom: 10
  },
  chevron: {
    alignSelf: 'center',
    marginLeft: 10,
    color: palette.neutralTertiary,
    fontSize: fonts.large.fontSize,
    flexShrink: 0
  }
});
const controlStyles = {
  root: {
    margin: '0 30px 20px 0',
    maxWidth: '300px'
  }
};
const dropdownStyles: Partial<IDropdownStyles> = {
  dropdown: { width: 300 }
};


let appContext;



export class DropdownControlledMulti extends React.Component<IDropdownControlledMultiProps, IDropdownControlledMultiState> {
  constructor(props) {
    super(props);

    this.state = {
      selectedItems: []
    };
  }

  public render() {
    const { selectedItems } = this.state;
    let choiceObjects = [];
    this.props.choices.map(choice => {
      choiceObjects.push({ key: choice.split(' ').join(''), text: choice });
    });
    return (
      <Dropdown
        placeholder={this.props.placeholder}
        label={this.props.label}
        selectedKeys={selectedItems}
        onChange={this._onChange}
        multiSelect
        options={choiceObjects}
        styles={{ dropdown: { width: 300 } }}
      />
    );
  }

  private _onChange = (event: React.FormEvent<HTMLDivElement>, item: IDropdownOption): void => {
    const newSelectedItems = [...this.state.selectedItems];

    if (item.selected) {
      newSelectedItems.push(item.key as string);
    } else {
      const currIndex = newSelectedItems.indexOf(item.key as string);
      if (currIndex > -1) {
        newSelectedItems.splice(currIndex, 1);
      }
    }
    this.setState({
      selectedItems: newSelectedItems
    },
    );
  }
}

export class ContactCard extends React.Component<IContactCardProps, IContactCardState> {

  constructor(props) {
    super(props);

    this.state = {
      item: this.props.item,
      showOrganization: this.props.showOrganization,
      showDepartment: this.props.showDepartment,
      showDivision: this.props.showDivision
    };
  }

  // public componentDidMount() {
  //   console.groupCollapsed('ContactCard -> componentDidMount');
  //   console.log('props', this.props);
  //   console.log('state', this.state);
  //   console.groupEnd();
  // }

  public componentDidUpdate(previousProps: IContactCardProps, previousState: IContactCardState) {
    // console.groupCollapsed('ContactCard -> componentDidUpdate');
    // console.groupEnd();
    if (previousState.item != this.props.item) {
      this.setState({ item: this.props.item }, () => {
      });
    }
    if (previousState.showOrganization != this.props.showOrganization) {
      this.setState({ showOrganization: this.props.showOrganization }, () => {
      });
    }
    if (previousState.showDepartment != this.props.showDepartment) {
      this.setState({ showDepartment: this.props.showDepartment }, () => {
      });
    }
    if (previousState.showDivision != this.props.showDivision) {
      this.setState({ showDivision: this.props.showDivision }, () => {
      });
    }
  }

  public render() {
    const searchTerms = this.props.searchTerms;

    let highlightHits = (str) => {
      for (let term of searchTerms) {
        const searchTermRegex = new RegExp(term.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'), "ig");
        const searchTermHighlighted = '<span style="background-color:yellow;">$&</span>';
        str = str.replace(searchTermRegex, searchTermHighlighted);
      }
      return str;
    };

    return (

      <div
        key={this.props.item.Id}
        className={this.props.size == 'large' ? styles.contactItem : [styles.contactItem, styles.small].join(' ')}
        data-item-id={this.props.item.Id}
      >
        <div className={styles.contactItemImg}>
          <Link href={"https://delve-gcc.office.com/?p=" + this.props.item.Email + "&v=work"} target="about:blank">
            <PersonaCoin
              text={this.props.item.FirstName != null ? this.props.item.FirstName + ' ' + this.props.item.Title : this.props.item.Title}
              coinSize={this.props.size == 'large' ? 100 : 50}
              showInitialsUntilImageLoads={true}
            // imageUrl='https://googlechrome.github.io/samples/picture-element/images/kitten-large.png'
            />
          </Link>
          {/* <FacepileBasicExample
            personas={[{
              personaName: this.props.item.FirstName != null ? this.props.item.FirstName + ' ' + this.props.item.Title : this.props.item.Title,
            }]}
            personaSize={this.props.size == 'large' ? 100 : 50}
          /> */}
        </div>
        <div className={styles.contactItemDetails}>

          <div className={styles.padBottom}>
            <Link href={"https://delve-gcc.office.com/?p=" + this.props.item.Email + "&v=work"} target="about:blank">
              <div className={[styles.contactItemFullName, styles.contactItemFieldBody].join(' ')}
                dangerouslySetInnerHTML={{
                  __html: highlightHits(this.props.item.FirstName != null ? this.props.item.FirstName + ' ' + this.props.item.Title : this.props.item.Title)
                }}
              />
            </Link>
            {this.props.item.JobTitle != null
              ? <div className={styles.contactItemFieldBody}
                dangerouslySetInnerHTML={{ __html: highlightHits(this.props.item.JobTitle) }} />
              : ''
            }
          </div>

          <div className={styles.padBottom}>
            {this.props.item.Organization != null && this.state.showOrganization
              ? <div className={styles.contactItemFieldBody}
                dangerouslySetInnerHTML={{ __html: highlightHits(this.props.item.Organization) }} />
              : ''
            }
            {this.props.item.Company != null && this.state.showDepartment
              ? <div>
                <span className={styles.contactItemFieldLabel}>Department: </span>
                <span className={styles.contactItemFieldBody}
                  dangerouslySetInnerHTML={{ __html: highlightHits(this.props.item.Company) }} />
              </div>
              : ''
            }
            {this.props.item.Division != null && this.state.showDivision
              ? <div>
                <span className={styles.contactItemFieldLabel}>Division: </span>
                <span className={styles.contactItemFieldBody}
                  dangerouslySetInnerHTML={{ __html: highlightHits(this.props.item.Division) }} />
              </div>
              : ''}
            {this.props.item.Program != null
              ? <div>
                <span className={styles.contactItemFieldLabel}>Program: </span>
                <span className={styles.contactItemFieldBody}
                  dangerouslySetInnerHTML={{ __html: highlightHits(this.props.item.Program) }} />
              </div>
              : ''}
          </div>
          {this.props.item.Email != null
            ? <div className={styles.contactItemFieldBody}>
              <a href={'mailto:' + this.props.item.Email}>
                {this.props.item.Email}
              </a>
            </div>
            : ''
          }
          {this.props.item.WorkPhone != null
            ? <div className={styles.contactItemFieldBody}>{this.props.item.WorkPhone}</div>
            : ''
          }
          {this.props.item.WorkAddress != null
            ? <div className={styles.contactItemFieldBody}>{this.props.item.WorkAddress}</div>
            : ''
          }
        </div>

      </div>

    );
  }

}

export class ContactCardGrid extends React.Component<IContactCardGridProps, IContactCardGridState> {

  constructor(props) {
    super(props);

    this.state = {
      items: this.props.items,
      showOrganization: this.props.showOrganization,
      showDepartment: this.props.showDepartment,
      showDivision: this.props.showDivision
    };
  }

  public render() {
    return (
      <div>
        {this.state.items.map((item) => {
          return (
            <ContactCard
              item={item}
              searchTerms={this.props.searchTerms}
              size={this.props.size}
              showOrganization={this.props.showOrganization}
              showDepartment={this.props.showDepartment}
              showDivision={this.props.showDivision}
            />
          );
        })}
      </div>
    );
  }

  public componentDidMount() {
    console.groupCollapsed('ContactCardGrid -> componentDidMount');
    console.log('props', this.props);
    console.log('state', this.state);
    console.groupEnd();
  }

  public componentDidUpdate(previousProps: IContactCardGridProps, previousState: IContactCardGridState) {
    console.groupCollapsed('ContactCardGrid -> componentDidUpdate');
    console.log('previousProps', previousProps);
    console.log('props', this.props);
    console.log('previousState', previousState);
    console.log('state', this.state);
    console.groupEnd();
    if (previousState.items != this.props.items) {
      this.setState({ items: this.props.items }, () => {
      });
    }
  }

  private _getKey(item: any, index?: number): string {
    return item.key;
  }

}

export class DetailsListCustomColumnsResults extends React.Component<IDetailsListCustomColumnsResultsProp, IDetailsListCustomColumnsResultsState> {

  constructor(props) {
    super(props);

    this.state = {
      sortedItems: this.props.items,
      columns: _buildColumns(this.props.items, this.props.showOrganization, this.props.showDepartment, this.props.showDivision),
      searchTerms: this.props.searchTerms,
      order: this.props.order,
      showOrganization: this.props.showOrganization,
      showDepartment: this.props.showDepartment,
      showDivision: this.props.showDivision
    };

    this._renderItemColumn = this._renderItemColumn.bind(this);
  }

  public render() {
    const { sortedItems, columns, searchTerms } = this.state;
    columns.map(column => {
      column.isResizable = true;
      column.name = column.fieldName == 'Company' ? 'Department'
        : column.fieldName == 'Title' ? 'Last Name'
          : column.fieldName.replace(/([A-Z])/g, ' $1').trim();
    });
    console.log('DetailsListCustomColumnsResults -> render -> columns', columns);

    return (
      <ShimmeredDetailsList
        items={sortedItems}
        setKey="set"
        columns={columns}
        onRenderItemColumn={this._renderItemColumn}
        onColumnHeaderClick={this._onColumnClick}
        onItemInvoked={this._onItemInvoked}
        onColumnHeaderContextMenu={this._onColumnHeaderContextMenu}
        ariaLabelForSelectionColumn="Toggle selection"
        ariaLabelForSelectAllCheckbox="Toggle selection for all items"
        checkButtonAriaLabel="Row checkbox"
        searchTerms={searchTerms}
      />
    );
  }

  public sendData = (order) => {
    this.props.parentCallback(order);
  }

  public componentDidMount() {
    console.groupCollapsed('DetailsListCustomColumnsResults -> componentDidMount');
    console.log('props', this.props);
    console.log('state', this.state);
    console.groupEnd();
  }

  public componentDidUpdate(previousProps: IDetailsListCustomColumnsResultsProp, previousState: IDetailsListCustomColumnsResultsState) {
    console.groupCollapsed('DetailsListCustomColumnsResults -> componentDidUpdate');
    console.log('previousProps', previousProps);
    console.log('props', this.props);
    console.log('previousState', previousState);
    console.log('state', this.state);
    console.groupEnd();
    if (previousState.sortedItems != this.props.items) {
      this.setState({ sortedItems: this.props.items }, () => {
      });
    }
    if (previousState.order != this.props.order) {
      this.setState({ order: this.props.order }, () => {
      });
    }
    if (previousState.showOrganization != this.props.showOrganization) {
      this.setState({
        showOrganization: this.props.showOrganization,
        columns: _buildColumns(this.props.items, this.props.showOrganization, this.props.showDepartment, this.props.showDivision)
      }, () => {

      });
    }
    if (previousState.showDepartment != this.props.showDepartment) {
      this.setState({
        showDepartment: this.props.showDepartment,
        columns: _buildColumns(this.props.items, this.props.showOrganization, this.props.showDepartment, this.props.showDivision)
      }, () => {
      });
    }
    if (previousState.showDivision != this.props.showDivision) {
      this.setState({
        showDivision: this.props.showDivision,
        columns: _buildColumns(this.props.items, this.props.showOrganization, this.props.showDepartment, this.props.showDivision)
      }, () => {
      });
    }
  }


  public _renderItemColumn(item: IResult, index: number, column: IColumn, searchTerms: any) {
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
    const { columns } = this.state;
    let { sortedItems } = this.state;
    let isSortedDescending = column.isSortedDescending;

    if (column.isSorted) {
      isSortedDescending = !isSortedDescending;
    }

    this.setState({
      order: column.fieldName
    }, () => {
      this.sendData(this.state.order);
    });
  }

  private _onColumnHeaderContextMenu(column: IColumn | undefined, ev: React.MouseEvent<HTMLElement> | undefined): void {
    console.log(`column ${column!.key} contextmenu opened.`);
  }

  private _onItemInvoked(item: any, index: number | undefined): void {
    alert(`Item ${item.name} at index ${index} has been invoked.`);
  }
}

export class FilterPanel extends React.Component<IFilterPanelProps, IFilterPanelState> {
  constructor(props) {
    super(props);

    this.state = {
      showPanel: this.props.showPanel,
      hasChoiceData: false,
      filters: '',
      filtersOrganization: [],
      filtersDepartment: [],
      filtersDivision: [],
      clearFilters: false
    };
  }

  public componentDidMount() {
    console.groupCollapsed('FilterPanel -> componentDidMount');
    console.log('props', this.props);
    console.log('state', this.state);
    console.groupEnd();
  }

  public componentDidUpdate(previousProps: IFilterPanelProps, previousState: IFilterPanelState) {
    console.groupCollapsed('FilterPanel -> componentDidUpdate');
    console.log('previousProps', previousProps);
    console.log('props', this.props);
    console.log('previousState', previousState);
    console.log('state', this.state);
    console.groupEnd();
    if (previousState.showPanel != this.props.showPanel) {
      this.setState({ showPanel: this.props.showPanel }, () => {
        this.sendData(this.state.showPanel, this.state.filters, this.state.filtersOrganization.length, this.state.filtersDepartment.length, this.state.filtersDivision.length, this.state.clearFilters);
      });
    }

    if (previousState.hasChoiceData === false && this.state.hasChoiceData === false) {
      this.setState({ hasChoiceData: true }, () => {
        this.getRESTResults();
      });
    }

    if (previousState.clearFilters != this.props.clearFilters) {
      this.setState({
        clearFilters: this.props.clearFilters/* , needUpdate: true */
      });
    }

  }

  public availOrganizations = [];
  public availOrganizationsObject = [];
  public availDepartments = [];
  public availDepartmentsObject = [];
  public availDivisions = [];
  public availDivisionsObject = [];

  public sortDropdowns(a, b) {
    return /* (a, b) =>  */(a.text > b.text) ? 1 : -1;
  }

  public getRESTResults() {
    const myPromise = new Promise((resolve, reject) => {
      const searchSourceUrl = "https://auroragov.sharepoint.com/sites/PhoneList";
      const listName = "EmployeeContactList";
      const select = "$select=Company,JobTitle,Division,Program,Organization";
      const top = "$top=5000";
      const requestUrl = searchSourceUrl + "/_api/web/lists/getbytitle('" + listName + "')/items?" + select + "&" + top;
      appContext.spHttpClient.get(requestUrl, SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) => {
          if (response.ok) {
            response.json().then((responseJSON) => {
              if (responseJSON != null) {
                let items: any[] = responseJSON.value;
                resolve(items);
              }
              reject(new Error('Something went wrong.'));
            });
          }
        });
    });
    const onResolved = (items) => {
      items.map(item => {

        if (item.Organization != null) {
          if (this.availOrganizations.indexOf(item.Organization) === -1) {
            this.availOrganizations.push(item.Organization);
            this.availOrganizationsObject.push({
              key: item.Organization.split(' ').join(''),
              text: item.Organization
            });
          }
          this.availOrganizationsObject.sort(this.sortDropdowns);
        }

        if (item.Company != null) {
          if (this.availDepartments.indexOf(item.Company) === -1) {
            this.availDepartments.push(item.Company);
            this.availDepartmentsObject.push({
              key: item.Company.split(' ').join(''),
              text: item.Company
            });
          }
          this.availDepartmentsObject.sort(this.sortDropdowns);
        }

        if (item.Division != null) {
          if (this.availDivisions.indexOf(item.Division) === -1) {
            this.availDivisions.push(item.Division);
            this.availDivisionsObject.push({
              key: item.Division.split(' ').join(''),
              text: item.Division
            });
          }
          this.availDivisionsObject.sort(this.sortDropdowns);
        }

      });
    };
    const onRejected = (error) => console.log(error);

    myPromise.then(onResolved, onRejected);

  }

  public sendData = (showPanel, filters, hasFiltersOrganization, hasFiltersDepartment, hasFiltersDivision, clearFilters) => {
    console.groupCollapsed('FilterPanel -> sendData');
    console.log('showPanel, filters, hasFiltersOrganization, hasFiltersDepartment, hasFiltersDivision, clearFilters', showPanel, filters, hasFiltersOrganization, hasFiltersDepartment, hasFiltersDivision, clearFilters);
    console.log('state', this.state);
    console.groupEnd();
    this.props.parentCallback(showPanel, filters, hasFiltersOrganization, hasFiltersDepartment, hasFiltersDivision, clearFilters);
  }

  private _showPanel = (): void => {
    this.setState({ showPanel: true });
  }

  private _hidePanel = (): void => {
    this.setState(
      { showPanel: false },
      () => { this.sendData(this.state.showPanel, this.state.filters, this.state.filtersOrganization.length, this.state.filtersDepartment.length, this.state.filtersDivision.length, this.state.clearFilters); }
    );
  }
 
  // private _onDismiss = (ev?: React.SyntheticEvent<HTMLElement>) => {
  //   if (!ev) {
  //     console.log('Panel dismissed.');
  //     return;
  //   }

  //   console.log('Close button clicked or light dismissed.');
  //   if (ev.nativeEvent.srcElement && (ev.nativeEvent.srcElement as Element).className.indexOf('ms-Button-icon') !== -1) {
  //     console.log('Close button clicked.');
  //   }
  //   if (ev.nativeEvent.srcElement && (ev.nativeEvent.srcElement as Element).className.indexOf('ms-Overlay') !== -1) {
  //     console.log('Light dismissed.');
  //   }
  //   this._hidePanel();
  //   // this.sendData(false);
  // }

  private _applyFilters = () => {

    let restFilters = [];
    let hasFiltersOrganization = false;
    let hasFiltersDepartment = false;
    let hasFiltersDivision = false;
    console.groupCollapsed('FilterPanel -> _applyFilters');
    console.log('this.state.filtersOrganization', this.state.filtersOrganization, this.state.filtersOrganization.length);
    console.log('this.state.filtersDepartment', this.state.filtersDepartment, this.state.filtersDepartment.length);
    console.log('this.state.filtersDivision', this.state.filtersDivision, this.state.filtersDivision.length);
    console.groupEnd();
    if (this.state.filtersOrganization.length) {
      const restFiltersOrganization = "(Organization eq '" + this.state.filtersOrganization.join("' or Organization eq '") + "')";
      restFilters.push(restFiltersOrganization);
      hasFiltersOrganization = true;
    }
    if (this.state.filtersDepartment.length) {
      const restFiltersDepartment = "(Company eq '" + this.state.filtersDepartment.join("' or Company eq '") + "')";
      restFilters.push(restFiltersDepartment);
      hasFiltersDepartment = true;
    }
    if (this.state.filtersDivision.length) {
      const restFiltersDivision = "(Division eq '" + this.state.filtersDivision.join("' or Division eq '") + "')";
      restFilters.push(restFiltersDivision);
      hasFiltersDivision = true;
    }

    this.setState(
      { filters: restFilters.join(' and ') },
      () => {
        this.sendData(this.state.showPanel, this.state.filters, hasFiltersOrganization, hasFiltersDepartment, hasFiltersDivision, this.state.clearFilters);
      }
    );
  }

  private _clearFilters = () => {
    console.log('_clearFilters');
    this.setState(
      {
        showPanel: false,
        filters: '',
        filtersOrganization: [],
        filtersDepartment: [],
        filtersDivision: [],
        clearFilters: true
      },
      () => {
        this.sendData(this.state.showPanel, this.state.filters, false, false, false, true);
      }
    );
  }

  private _onRenderFooterContent = () => {
    const applyFilterIcon: IIconProps = { iconName: 'WaitlistConfirmMirrored' };
    const hideFilterIcon: IIconProps = { iconName: 'Hide' };
    const clearFilterIcon: IIconProps = { iconName: 'ClearFilter' };
    return (
      <div>
        <DefaultButton
          iconProps={applyFilterIcon}
          onClick={this._applyFilters}
        >Apply</DefaultButton>
        <DefaultButton
          iconProps={hideFilterIcon}
          styles={{ root: { marginLeft: 15 } }}
          onClick={this._hidePanel}
        >Hide</DefaultButton>
        <DefaultButton
          iconProps={clearFilterIcon}
          styles={{ root: { marginLeft: 15 } }}
          onClick={this._clearFilters}
        >Clear</DefaultButton>
      </div>
    );
  }

  private _onFilterChangeOrganization = (e) => {
    if (e.target.checked) {
      let newFilters = this.state.filtersOrganization;
      newFilters.push(e.target.title);
      this.setState({
        filtersOrganization: newFilters
      },
        () => {
        }
      );
    }
  }

  private _onFilterChangeDepartment = (e) => {
    if (e.target.checked) {
      let newFilters = this.state.filtersDepartment;
      newFilters.push(e.target.title);
      this.setState({
        filtersDepartment: newFilters
      },
        () => {
        }
      );
    }
  }

  private _onFilterChangeDivision = (e) => {
    if (e.target.checked) {
      let newFilters = this.state.filtersDivision;
      newFilters.push(e.target.title);
      this.setState({
        filtersDivision: newFilters
      },
        () => {
        }
      );
    }
  }

  public render() {
    return (
      <Panel
        key={this.state.clearFilters ? 'ReRender' : 'noReRender'}
        isOpen={this.state.showPanel}
        closeButtonAriaLabel='Close'
        isLightDismiss={true}
        headerText='Light Dismiss Panel'
        onDismiss={this._hidePanel}
        onRenderFooterContent={this._onRenderFooterContent}
        isHiddenOnDismiss={true}
        isFooterAtBottom={true}
        type={PanelType.custom}
        customWidth='420px'
      >
        <Dropdown
          placeholder='Select departments...'
          label='Department'
          onChange={this._onFilterChangeDepartment}
          multiSelect
          options={this.availDepartmentsObject}
          styles={{ dropdown: { width: 300 } }}
        />
        <Dropdown
          placeholder='Select divisions...'
          label='Division'
          onChange={this._onFilterChangeDivision}
          multiSelect
          options={this.availDivisionsObject}
          styles={{ dropdown: { width: 300 } }}
        />
        <Dropdown
          placeholder='Select organizations...'
          label='Organization'
          onChange={this._onFilterChangeOrganization}
          multiSelect
          options={this.availOrganizationsObject}
          styles={{ dropdown: { width: 300 } }}
        />
      </Panel>
    );
  }

}

export class CommandBarSearchControls extends React.Component<ICommandBarSearchControlsProps, ICommandBarSearchControlsState> {

  constructor(props) {
    super(props);

    this.state = {
      view: this.props.view,
      order: this.props.order,
      size: 'large',
      showPanel: this.props.showPanel,
      filters: this.props.filters,
      hasFiltersOrganization: false,
      hasFiltersDepartment: false,
      hasFiltersDivision: false,
      clearFilters: this.props.clearFilters
    };

    this.handleViewTilesClick = this.handleViewTilesClick.bind(this);
    this.handleViewListClick = this.handleViewListClick.bind(this);
    this.handleSortTilesClick = this.handleSortTilesClick.bind(this);
    this.handleFilterClick = this.handleFilterClick.bind(this);
  }

  public componentDidMount() {
    console.groupCollapsed('CommandBarSearchControls -> componentDidMount');
    console.log('props', this.props);
    console.log('state', this.state);
    console.groupEnd();
  }

  public componentDidUpdate(previousProps: ICommandBarSearchControlsProps, previousState: ICommandBarSearchControlsState) {
    console.groupCollapsed('CommandBarSearchControls -> componentDidUpdate');
    console.log('previousProps', previousProps);
    console.log('props', this.props);
    console.log('previousState', previousState);
    console.log('state', this.state);
    console.groupEnd();
    if (previousState.filters != this.props.filters) {
      this.setState({ filters: this.props.filters }, () => {
        this.sendData(true, this.state.view, this.state.order, this.state.size, this.state.showPanel, this.state.filters, this.state.hasFiltersOrganization, this.state.hasFiltersDepartment, this.state.hasFiltersDivision, false);
      });
    }
    if (previousState.clearFilters != this.props.clearFilters) {
      this.setState({
        clearFilters: this.props.clearFilters,
        showPanel: false
      }, () => {
        this.sendData(true, this.state.view, this.state.order, this.state.size, this.state.showPanel, this.state.filters, this.state.hasFiltersOrganization, this.state.hasFiltersDepartment, this.state.hasFiltersDivision, this.state.clearFilters);
      });
    }
  }

  public sendData = (boolVal, view, order, size, showPanel, filters, hasFiltersOrganization, hasFiltersDepartment, hasFiltersDivision, clearFilters) => {
    this.props.parentCallback(boolVal, view, order, size, showPanel, filters, hasFiltersOrganization, hasFiltersDepartment, hasFiltersDivision, clearFilters);
  }

  public handleFilterClick = () => {
    console.log('filter clicked');
    this.setState({
      showPanel: !this.state.showPanel
    });
  }

  public handleSortTilesClick = (orderClicked) => {
    console.log('order clicked');
    this.setState({
      order: orderClicked
    }, () => {
      this.sendData(true, this.state.view, this.state.order, this.state.size, this.state.showPanel, this.state.filters, this.state.hasFiltersOrganization, this.state.hasFiltersDepartment, this.state.hasFiltersDivision, false);
    });
  }

  public handleViewTilesClick = () => {
    console.log('Tiles');
    this.setState({
      view: 'Tiles'
    }, () => {
      this.sendData(true, this.state.view, this.state.order, this.state.size, this.state.showPanel, this.state.filters, this.state.hasFiltersOrganization, this.state.hasFiltersDepartment, this.state.hasFiltersDivision, false);
    });
  }

  public handleViewListClick = () => {
    console.log('List');
    this.setState({
      view: 'List'
    }, () => {
      this.sendData(true, this.state.view, this.state.order, this.state.size, this.state.showPanel, this.state.filters, this.state.hasFiltersOrganization, this.state.hasFiltersDepartment, this.state.hasFiltersDivision, false);
    });
  }

  public handleTileSizeClick = (sizeClicked) => {
    console.log('size clicked');
    this.setState({
      size: sizeClicked
    }, () => {
      this.sendData(true, this.state.view, this.state.order, this.state.size, this.state.showPanel, this.state.filters, this.state.hasFiltersOrganization, this.state.hasFiltersDepartment, this.state.hasFiltersDivision, false);
    });
  }

  public callbackFromFilterPanelToCommandBar = (showPanel, filters, hasFiltersOrganization, hasFiltersDepartment, hasFiltersDivision, clearFilters) => {
    this.setState({
      showPanel: showPanel,
      filters: filters,
      hasFiltersOrganization: hasFiltersOrganization,
      hasFiltersDepartment: hasFiltersDepartment,
      hasFiltersDivision: hasFiltersDivision,
      clearFilters: clearFilters
    },
      () => {
        console.groupCollapsed('CommandBarSearchControls -> callbackFromFilterPanelToCommandBar');
        console.log('showPanel, filters, hasFiltersOrganization, hasFiltersDepartment, hasFiltersDivision, clearFilters', showPanel, filters, hasFiltersOrganization, hasFiltersDepartment, hasFiltersDivision, clearFilters);
        console.log('this.state', this.state);
        console.groupEnd();
        this.sendData(true, this.state.view, this.state.order, this.state.size, this.state.showPanel, this.state.filters, this.state.hasFiltersOrganization, this.state.hasFiltersDepartment, this.state.hasFiltersDivision, clearFilters);
      }
    );
  }

  private getItems = () => {
    if (this.state.view == 'Tiles') {
      return [
        {
          key: 'size',
          name: 'Tile Size',
          ariaLabel: 'Tile Size',
          iconProps: {
            iconName: 'SizeLegacy'
          },
          onClick: () => { this.handleViewListClick(); },
          subMenuProps: {
            items: [
              {
                key: 'small',
                name: 'Small',
                iconProps: {
                  iconName: 'GridViewSmall'
                },
                onClick: () => {
                  this.handleTileSizeClick('small');
                }
              },
              {
                key: 'large',
                name: 'Large',
                iconProps: {
                  iconName: 'GridViewMedium'
                },
                onClick: () => {
                  this.handleTileSizeClick('large');
                }
              }
            ]
          }

        },
        {
          key: 'list',
          name: 'Switch to List View',
          ariaLabel: 'Switch to List View',
          iconProps: {
            iconName: 'GroupedList'
          },
          onClick: () => { this.handleViewListClick(); }
        }
      ];
    }
    return [
      {
        key: 'tile',
        name: 'Switch to Grid View',
        ariaLabel: 'Switch to Grid View',
        iconProps: {
          iconName: 'Tiles'
        },
        onClick: () => { this.handleViewTilesClick(); }
      }
    ];
  }

  // private getOverlflowItems = () => {
  //   return [
  //     {
  //       key: 'move',
  //       name: 'Move to...',
  //       onClick: () => console.log('Move to'),
  //       iconProps: {
  //         iconName: 'MoveToFolder'
  //       }
  //     },
  //     {
  //       key: 'copy',
  //       name: 'Copy to...',
  //       onClick: () => console.log('Copy to'),
  //       iconProps: {
  //         iconName: 'Copy'
  //       }
  //     },
  //     {
  //       key: 'rename',
  //       name: 'Rename...',
  //       onClick: () => console.log('Rename'),
  //       iconProps: {
  //         iconName: 'Edit'
  //       }
  //     }
  //   ];
  // }

  private getFarItems = () => {
    if (this.state.view == 'Tiles') {
      return [
        {
          key: 'sort',
          name: 'Sort',
          ariaLabel: 'Sort',
          iconProps: {
            iconName: 'SortLines'
          },
          subMenuProps: {
            items: [
              {
                key: 'firstName',
                name: 'First Name',
                iconProps: {
                  iconName: 'UserOptional'
                },
                onClick: () => {
                  this.handleSortTilesClick('FirstName');
                }
              },
              {
                key: 'lastName',
                name: 'Last Name',
                iconProps: {
                  iconName: 'UserOptional'
                },
                onClick: () => {
                  this.handleSortTilesClick('Title');
                }
              },
              {
                key: 'organization',
                name: 'Organization',
                iconProps: {
                  iconName: 'Org'
                },
                onClick: () => {
                  this.handleSortTilesClick('Organization');
                }
              },
              {
                key: 'department',
                name: 'Department',
                iconProps: {
                  iconName: 'Teamwork'
                },
                onClick: () => {
                  this.handleSortTilesClick('Company');
                }
              }
            ]
          }
        },
        {
          key: 'filter',
          name: 'Filter',
          ariaLabel: 'Filter',
          iconProps: {
            iconName: 'Filter'
          },
          onClick: () => {
            this.handleFilterClick();
          }
        }
      ];
    }
    return [
      {
        key: 'filter',
        name: 'Filter',
        ariaLabel: 'Filter',
        iconProps: {
          iconName: 'Filter'
        },
        onClick: () => {
          this.handleFilterClick();
        }
      }
    ];
  }

  public render(): JSX.Element {
    return (<div>
      <CommandBar
        items={this.getItems()}
        farItems={this.getFarItems()}
        ariaLabel={'Use left and right arrow keys to navigate between commands'}
      />
      <FilterPanel
        parentCallback={this.callbackFromFilterPanelToCommandBar}
        showPanel={this.state.showPanel}
        filters={this.state.filters}
        clearFilters={this.state.clearFilters}
      />
    </div>);
  }

}

export class ContactSearchBox extends React.Component<IContactSearchBoxProps, IContactSearchBoxState> {

  constructor(props) {
    super(props);

    this.state = {
      searchTerms: [],
      items: [],
      view: this.props.view,
      order: this.props.order,
      needUpdate: false,
      showPanel: false,
      filters: this.props.filters,
      hasFiltersOrganization: this.props.hasFiltersOrganization,
      hasFiltersDepartment: this.props.hasFiltersDepartment,
      hasFiltersDivision: this.props.hasFiltersDivision,
      showOrganization: this.props.showOrganization,
      showDepartment: this.props.showDepartment,
      showDivision: this.props.showDivision,
      clearFilters: this.props.clearFilters
    };

    this.handleChange = this.handleChange.bind(this);
    this.handleClear = this.handleClear.bind(this);
  }

  public componentDidMount() {
    console.groupCollapsed('ContactSearchBox -> componentDidMount');
    console.log('props', this.props);
    console.log('state', this.state);
    console.groupEnd();
  }

  public componentDidUpdate(previousProps: IContactSearchBoxProps, previousState: IContactSearchBoxState) {
    console.groupCollapsed('ContactSearchBox -> componentDidUpdate');
    console.log('previousProps', previousProps);
    console.log('props', this.props);
    console.log('previousState', previousState);
    console.log('state', this.state);
    console.groupEnd();
    if (previousState.order != this.props.order) {
      this.setState({ order: this.props.order, needUpdate: true }, () => {
        if (this.state.view == 'List') {
          this.getRESTResults(this.state.searchTerms);
        }
      });
    }
    if (previousState.size != this.props.size) {
      this.setState({ size: this.props.size, needUpdate: true }, () => {
      });
    }
    if (previousState.showPanel != this.props.showPanel) {
      this.setState({ showPanel: this.props.showPanel, needUpdate: true }, () => {
      });
    }
    if (previousState.filters != this.state.filters) {
      this.getRESTResults(this.state.searchTerms);
    }
    if (previousState.showOrganization != this.props.showOrganization) {
      this.setState({ showOrganization: this.props.showOrganization, needUpdate: true }, () => {
        this.getRESTResults(this.state.searchTerms);
      });
    }
    if (previousState.showDepartment != this.props.showDepartment) {
      this.setState({ showDepartment: this.props.showDepartment, needUpdate: true }, () => {
        this.getRESTResults(this.state.searchTerms);
      });
    }
    if (previousState.showDivision != this.props.showDivision) {
      this.setState({ showDivision: this.props.showDivision, needUpdate: true }, () => {
        this.getRESTResults(this.state.searchTerms);
      });
    }
    if (previousState.clearFilters != this.props.clearFilters) {
      this.setState({
        clearFilters: this.props.clearFilters/* , needUpdate: true */
      });
    }


  }

  public sendData = (boolVal, childData, searchTerms, view, order, size, showPanel, filters, hasFiltersOrganization, hasFiltersDepartment, hasFiltersDivision, clearFilters) => {
    this.props.parentCallback(boolVal, childData, searchTerms, view, order, size, showPanel, filters, hasFiltersOrganization, hasFiltersDepartment, hasFiltersDivision, clearFilters);
  }

  public handleChange = debounce(e => {
    if (e.length) {
      this.getRESTResults(e);
    }
    else {
      console.log('no data yet');
    }
  }, 1000);

  public getRESTResults(e) {
    console.groupCollapsed('ContactSearchBox -> getRESTResults');
    let searchTerms = [];
    const myPromise = new Promise((resolve, reject) => {
      if (e.constructor === Array) {
        searchTerms = e;
      }
      else {
        searchTerms = e.split(' ');
      }
      let searchFilters = [];
      const searchFields = [
        'Title',
        'FirstName',
        'JobTitle',
        'Program'
      ];
      console.log('this.props.hasFiltersOrganization', this.state.hasFiltersOrganization);
      console.log('this.props.hasFiltersDepartment', this.state.hasFiltersDepartment);
      console.log('this.props.hasFiltersDivision', this.state.hasFiltersDivision);
      if (!this.state.hasFiltersOrganization && this.state.showOrganization) {
        console.log('no org in refiners, add to searchFields');
        searchFields.push('Organization');
      }
      if (!this.state.hasFiltersDepartment && this.state.showDepartment) {
        console.log('no dept in refiners, add to searchFields');
        searchFields.push('Company');
      }
      if (!this.state.hasFiltersDivision && this.state.showDivision) {
        console.log('no div in refiners, add to searchFields');
        searchFields.push('Division');
      }
      for (let term of searchTerms) {
        for (let field of searchFields) {
          searchFilters.push("substringof('" + term + "'," + field + ")");
        }
      }
      const searchSourceUrl = "https://auroragov.sharepoint.com/sites/PhoneList";
      const listName = "EmployeeContactList";
      const select = "$select=Id,Title,FirstName,Email,Company,JobTitle,WorkPhone,WorkAddress,Division,Program,Organization,CellPhone";
      const top = "$top=100";

      const searchBarFilters = "(" + searchFilters.join(' or ') + ")";
      console.log('searchBarFilters', searchBarFilters);

      const refiners = this.state.filters != null && this.state.filters.length ? this.state.filters + " and " : '';
      console.log('refiners', refiners);

      const filter = "$filter=" + refiners + searchBarFilters;
      console.log('filter', filter);
      const sortOrder = '$orderby=' + this.state.order;
      const requestUrl = searchSourceUrl + "/_api/web/lists/getbytitle('" + listName + "')/items?" + select + "&" + top + "&" + filter + "&" + sortOrder;
      console.log('requestUrl', requestUrl);
      appContext.spHttpClient.get(requestUrl, SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) => {
          if (response.ok) {
            response.json().then((responseJSON) => {
              if (responseJSON != null) {
                let items: any[] = responseJSON.value;
                resolve(items);
              }
              reject(new Error('Something went wrong.'));
            });
          }
        });
    });
    const onResolved = (items) => {

      this.setState({
        items: items,
        searchTerms: searchTerms,
        view: this.props.view,
        order: this.props.order,
        size: this.props.size,
        hasFiltersOrganization: this.props.hasFiltersOrganization,
        hasFiltersDepartment: this.props.hasFiltersDepartment,
        hasFiltersDivision: this.props.hasFiltersDivision
      }, () => {
        this.sendData(true, this.state.items, this.state.searchTerms, this.state.view, this.state.order, this.state.size, this.state.showPanel, this.state.filters, this.state.hasFiltersOrganization, this.state.hasFiltersDepartment, this.state.hasFiltersDivision, false);
      });
    };
    const onRejected = (error) => console.log(error);

    myPromise.then(onResolved, onRejected);

    console.groupEnd();
  }

  public handleClear(e) {
    this.setState({
      items: [],
      searchTerms: '',
      order: ''
    },
      () => {
        this.sendData(true, this.state.items, this.state.searchTerms, this.state.view, this.state.order, this.state.size, this.state.showPanel, this.state.filters, this.state.hasFiltersOrganization, this.state.hasFiltersDepartment, this.state.hasFiltersDivision, false);
      }
    );
  }

  public callbackFromCommandBarToSearchBox = (boolVal, view, order, size, showPanel, filters, hasFiltersOrganization, hasFiltersDepartment, hasFiltersDivision, clearFilters) => {
    this.setState({
      view: view,
      order: order,
      needUpdate: boolVal,
      size: size,
      showPanel: showPanel,
      filters: filters,
      hasFiltersOrganization: hasFiltersOrganization,
      hasFiltersDepartment: hasFiltersDepartment,
      hasFiltersDivision: hasFiltersDivision,
      clearFilters: clearFilters
    },
      () => {
        this.sendData(true, this.state.items, this.state.searchTerms, this.state.view, this.state.order, this.state.size, this.state.showPanel, this.state.filters, this.state.hasFiltersOrganization, this.state.hasFiltersDepartment, this.state.hasFiltersDivision, clearFilters);
        this.handleChange(this.state.searchTerms);
      }
    );
  }

  public render() {
    const controls = this.state.items.length
      ? <CommandBarSearchControls
        parentCallback={this.callbackFromCommandBarToSearchBox}
        view={this.state.view}
        order={this.state.order}
        showPanel={this.state.showPanel}
        filters={this.state.filters}
        clearFilters={this.state.clearFilters}
      />
      : '';
    return (<div>
      <SearchBox
        underlined
        placeholder={this.props.searchBoxPlaceholder}
        onChange={this.handleChange}
        onClear={this.handleClear}
      />
      {controls}
    </div>);
  }

}

export class MainApp extends React.Component<IMainAppProps, IMainAppState> {

  constructor(props) {

    super(props);

    this.state = {
      needUpdate: false,
      items: [],
      searchTerms: '',
      view: 'Tiles',
      order: 'FirstName',
      size: 'large',
      showPanel: false,
      filters: '',
      hasFiltersOrganization: false,
      hasFiltersDepartment: false,
      hasFiltersDivision: false,
      clearFilters: false
    };

    this.callbackFromSearchBoxToMainApp = this.callbackFromSearchBoxToMainApp.bind(this);
  }

  public componentDidMount() {
    console.group('MainApp -> componentDidMount');
    console.log('this.props', this.props);
    console.log('this.state', this.state);
    console.groupEnd();
    console.log('asdfasdfasdfasdf', graph.users.getById('lhibbs@auroragov.org').photo.toUrl());
    console.log('asdfasdfasdfasdf', graph.users.getById('lhibbs@auroragov.org').photo.toUrlAndQuery());
  }

  public componentDidUpdate(previousProps: IMainAppProps, previousState: IMainAppState) {
    console.group('MainApp -> componentDidUpdate');
    console.log('previousProps', previousProps);
    console.log('props', this.props);
    console.log('previousState', previousState);
    console.log('state', this.state);
    console.groupEnd();
  }

  public componentWillUnmount() {
  }

  public callbackFromSearchBoxToMainApp = (boolVal, childData, searchTerms, view, order, size, showPanel, filters, hasFiltersOrganization, hasFiltersDepartment, hasFiltersDivision, clearFilters) => {
    this.setState({
      needUpdate: boolVal,
      items: childData,
      searchTerms: searchTerms,
      view: view,
      order: order,
      size: size,
      showPanel: showPanel,
      filters: filters,
      hasFiltersOrganization: hasFiltersOrganization,
      hasFiltersDepartment: hasFiltersDepartment,
      hasFiltersDivision: hasFiltersDivision,
      clearFilters: clearFilters
    },
      () => {
      }
    );
  }

  public callbackFromDetailsListToMainApp = (order) => {
    this.setState({
      order: order
    },
      () => {
      }
    );
  }

  public render() {
    const resultViewElement =
      this.state.searchTerms.length ?
        this.state.items.length ?
          this.state.view == 'Tiles'
            ? <ContactCardGrid
              items={this.state.items}
              searchTerms={this.state.searchTerms}
              size={this.state.size}
              showOrganization={this.props.showOrganization}
              showDepartment={this.props.showDepartment}
              showDivision={this.props.showDivision}
            />
            : <DetailsListCustomColumnsResults
              parentCallback={this.callbackFromDetailsListToMainApp}
              items={this.state.items}
              searchTerms={this.state.searchTerms}
              order={this.state.order}
              showOrganization={this.props.showOrganization}
              showDepartment={this.props.showDepartment}
              showDivision={this.props.showDivision}
            />
          : <div>{this.props.noResultText}</div>
        : <div>{this.props.initialResultText}</div>;

    return (<div id="appRootWrap">
      <h1>{this.props.appHeading}</h1>
      <ContactSearchBox
        parentCallback={this.callbackFromSearchBoxToMainApp}
        view={this.state.view}
        order={this.state.order}
        size={this.state.size}
        showPanel={this.state.showPanel}
        filters={this.state.filters}
        hasFiltersOrganization={this.state.hasFiltersOrganization}
        hasFiltersDepartment={this.state.hasFiltersDepartment}
        hasFiltersDivision={this.state.hasFiltersDivision}
        searchBoxPlaceholder={this.props.searchBoxPlaceholder}
        showOrganization={this.props.showOrganization}
        showDepartment={this.props.showDepartment}
        showDivision={this.props.showDivision}
        clearFilters={this.state.clearFilters}
      />
      {resultViewElement}
    </div>);
  }

}

export default class PhoneListSearchWebPart extends BaseClientSideWebPart<IPhoneListSearchWebPartProps> {

  public render(): void {

    appContext = this.context;

    const element = <div>
      <MainApp
        searchBoxPlaceholder={this.properties.searchBoxPlaceholder}
        appHeading={this.properties.appHeading}
        initialResultText={this.properties.initialResultText}
        noResultText={this.properties.noResultText}
        showOrganization={this.properties.showOrganization}
        showDepartment={this.properties.showDepartment}
        showDivision={this.properties.showDivision}
      />
    </div>;

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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
              groupName: 'Page Text',
              groupFields: [
                PropertyPaneTextField('appHeading', {
                  label: 'Heading',
                  description: 'The heading that shows above the search box.'
                }),
                PropertyPaneTextField('searchBoxPlaceholder', {
                  label: 'Search Box Placeholder Text',
                  description: 'Text that shows inside the search box before the user enters text.'
                }),
                PropertyPaneTextField('initialResultText', {
                  label: 'Initial Result Text',
                  description: 'Text that shows in the results pane before the user searches.',
                  multiline: true
                }),
                PropertyPaneTextField('noResultText', {
                  label: 'No Result Text',
                  description: 'Text that shows in the results pane if no results are found.',
                  multiline: true
                })
              ]
            },
            {
              groupName: 'Fields to Show',
              groupFields: [
                PropertyPaneCheckbox('showOrganization', {
                  text: 'Organization'
                }),
                PropertyPaneCheckbox('showDepartment', {
                  text: 'Department'
                }),
                PropertyPaneCheckbox('showDivision', {
                  text: 'Division'
                })
              ]
            }
          ]
        }
      ]
    };
  }

}




function _buildColumns(items: IResult[], showOrganization, showDepartment, showDivision): IColumn[] {
  console.log('_buildColumns -> items', items);

  let theColumns = [];

  items.map(item => {
    theColumns.push({
      FirstName: item.FirstName,
      Title: item.Title,
      JobTitle: item.JobTitle,
      ...showOrganization ? { Organization: item.Organization } : null,
      ...showDepartment ? { Company: item.Company } : null,
      ...showDivision ? { Division: item.Division } : null,
      Program: item.Program,
      Email: item.Email,
      WorkPhone: item.WorkPhone,
      WorkAddress: item.WorkAddress
    });
  });

  const columns = buildColumns(theColumns);

  return columns;
}
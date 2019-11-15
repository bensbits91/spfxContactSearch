var __extends = (this && this.__extends) || (function () {
    var extendStatics = Object.setPrototypeOf ||
        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { PropertyPaneTextField, PropertyPaneCheckbox, PropertyPaneDropdown } from '@microsoft/sp-property-pane';
import * as strings from 'PhoneListSearchWebPartStrings';
// import { Link } from 'office-ui-fabric-react/lib/Link';
import { update } from '@microsoft/sp-lodash-subset';
// import { debounce } from 'lodash';
import { /* sp,  */ Web } from '@pnp/sp';
import App from './components/App';
import './components/temp.css';
// let appContext;
var availOrganizationsObject = [];
var propPaneDepartments = [];
var propPaneDivisions = [];
var blankOption = {
    key: 'NoFilter',
    text: 'No Filter'
};
var PhoneListSearchWebPart = /** @class */ (function (_super) {
    __extends(PhoneListSearchWebPart, _super);
    function PhoneListSearchWebPart() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this.availOrganizations = [];
        return _this;
    }
    PhoneListSearchWebPart.prototype.onInit = function () {
        var _this = this;
        //  appContext = this.context;
        var theWeb = new Web(this.context.pageContext.web.absoluteUrl);
        var theList = theWeb.lists.getByTitle('EmployeeContactList');
        var select = 'Company,Division,Organization';
        this.getOptionsPromise = theList.items.select(select).getAll().then(function (items) {
            console.log('TCL: PhoneListSearchWebPart -> items', items);
            var departmentsTempArray = [];
            var divisionsTempArray = [];
            update(_this.properties, 'options_department', function () {
                return [];
            });
            update(_this.properties, 'options_division', function () {
                return [];
            });
            update(_this.properties, 'organizationOptions', function () {
                return [];
            });
            items.map(function (item) {
                if (item.Company != null) {
                    if (departmentsTempArray.indexOf(item.Company) === -1) {
                        departmentsTempArray.push(item.Company);
                        _this.properties.options_department.push({
                            key: item.Company.replace(/ /g, ''),
                            text: item.Company
                        });
                    }
                }
                if (item.Division != null) {
                    if (divisionsTempArray.indexOf(item.Division) === -1) {
                        divisionsTempArray.push(item.Division);
                        _this.properties.options_division.push({
                            key: item.Division.replace(/ /g, ''),
                            text: item.Division,
                            department: item.Company.replace(/ /g, '')
                        });
                    }
                }
                if (item.Organization != null) {
                    if (_this.availOrganizations.indexOf(item.Organization) === -1) {
                        _this.availOrganizations.push(item.Organization);
                        availOrganizationsObject.push({
                            key: item.Organization.replace(/ /g, ''),
                            text: item.Organization,
                            department: item.Company.replace(/ /g, ''),
                            division: item.Division.replace(/ /g, '')
                        });
                    }
                }
            });
            _this.properties.options_department.sort(_this.sortDropdowns);
            _this.properties.options_division.sort(_this.sortDropdowns);
            availOrganizationsObject.sort(_this.sortDropdowns);
            // const blankOption = {
            //    key: 'NoFilter',
            //    text: 'No Filter'
            // };
            propPaneDepartments = JSON.parse(JSON.stringify(_this.properties.options_department));
            propPaneDepartments.unshift(blankOption);
            propPaneDivisions = JSON.parse(JSON.stringify(_this.properties.options_division));
            propPaneDivisions.unshift(blankOption);
            //   this.render();
        }).then(function () {
            console.clear();
            console.log('%c PhoneListSearchWebPart -> this.properties', 'color:aqua', _this.properties);
            // console.log('%c PhoneListSearchWebPart -> propPaneDepartments', 'color:aqua', propPaneDepartments);
            // console.log('%c PhoneListSearchWebPart -> this.properties.options_division', 'color:magenta', this.properties.options_division);
            // console.log('%c PhoneListSearchWebPart -> propPaneDivisions', 'color:magenta', propPaneDivisions);
            // console.log('%c PhoneListSearchWebPart -> availOrganizationsObject', 'color:orange', availOrganizationsObject);
        });
        //  this.getOptionsPromise = this.getOptions();
        return this.getOptionsPromise;
    };
    PhoneListSearchWebPart.prototype.sortDropdowns = function (a, b) {
        return (a.key > b.key) ? 1 : -1;
    };
    //   private getOptions(): Promise<void> {
    //     return new Promise<void>((resolve2: (options) => void, reject2: (error: any) => void) => {
    //       const myPromise = new Promise((resolve, reject) => {
    //         const searchSourceUrl = "https://auroragov.sharepoint.com/sites/PhoneList";
    //         const listName = "EmployeeContactList";
    //         const select = "$select=Company,Division,Organization";
    //         const top = "$top=500";
    //         const requestUrl = searchSourceUrl + "/_api/web/lists/getbytitle('" + listName + "')/items?" + select + "&" + top;
    //         appContext.spHttpClient.get(requestUrl, SPHttpClient.configurations.v1)
    //           .then((response: SPHttpClientResponse) => {
    //             if (response.ok) {
    //               response.json().then((responseJSON) => {
    //                 if (responseJSON != null) {
    //                   let items: any[] = responseJSON.value;
    //                   resolve(items);
    //                 }
    //                 reject(new Error('Something went wrong.'));
    //               });
    //             }
    //           });
    //       });
    //       const onResolved = (items) => {
    //         let departmentsTempArray = [];
    //         let divisionsTempArray = [];
    //         update(this.properties, 'options_department', (): any => {
    //           return [];
    //         });
    //         update(this.properties, 'options_division', (): any => {
    //           return [];
    //         });
    //         update(this.properties, 'organizationOptions', (): any => {
    //           return [];
    //         });
    //         items.map(item => {
    //           if (item.Company != null) {
    //             if (departmentsTempArray.indexOf(item.Company) === -1) {
    //               departmentsTempArray.push(item.Company);
    //               this.properties.options_department.push({
    //                 key: item.Company.replace(/ /g, ''),
    //                 text: item.Company
    //               });
    //             }
    //           }
    //           if (item.Division != null) {
    //             if (divisionsTempArray.indexOf(item.Division) === -1) {
    //               divisionsTempArray.push(item.Division);
    //               this.properties.options_division.push({
    //                 key: item.Division.replace(/ /g, ''),
    //                 text: item.Division
    //               });
    //             }
    //           }
    //           if (item.Organization != null) {
    //             if (this.availOrganizations.indexOf(item.Organization) === -1) {
    //               this.availOrganizations.push(item.Organization);
    //               availOrganizationsObject.push({
    //                 key: item.Organization.replace(/ /g, ''),
    //                 text: item.Organization
    //               });
    //             }
    //           }
    //         });
    //         this.properties.options_department.sort(this.sortDropdowns);
    //         this.properties.options_division.sort(this.sortDropdowns);
    //         availOrganizationsObject.sort(this.sortDropdowns);
    //         const blankOption = {
    //           key: 'NoFilter',
    //           text: 'No Filter'
    //         };
    //         propPaneDepartments = JSON.parse(JSON.stringify(this.properties.options_department));
    //         propPaneDepartments.unshift(blankOption);
    //         propPaneDivisions = JSON.parse(JSON.stringify(this.properties.options_division));
    //         propPaneDivisions.unshift(blankOption);
    //         this.render();
    //       };
    //       const onRejected = (error) => { console.log(error); };
    //       myPromise.then(onResolved, onRejected);
    //       resolve2('good to go');
    //       reject2(new Error('Something went wrong.'));
    //     });
    //   }
    PhoneListSearchWebPart.prototype.render = function () {
        var _this = this;
        console.log('%c  PhoneListSearchWebPart -> this.properties', 'color:lime', this.properties);
        if (this.properties.options_department) {
            if (this.properties.prefilter_key_department) {
                if (this.properties.prefilter_key_department != 'NoFilter') {
                    var newDeparmentLabel_1 = this.properties.options_department.find(function (obj) { return obj.key == _this.properties.prefilter_key_department; }).text;
                    update(this.properties, 'prefilter_label_department', function () { return newDeparmentLabel_1; });
                    // update cascading division options
                    var newoptions_division = JSON.parse(JSON.stringify(this.properties.options_division)).filter(function (option) { return option.department == _this.properties.prefilter_key_department; });
                    // console.log('%c  PhoneListSearchWebPart -> this.properties.options_division', 'color:lime', this.properties.options_division);
                    // console.log('%c  PhoneListSearchWebPart -> newoptions_division', 'color:lime', newoptions_division);
                    // newoptions_division.filter(option => option.department == this.properties.prefilter_key_department);
                    // update(this.properties, 'options_division', (): any => { return newoptions_division; });
                    propPaneDivisions = JSON.parse(JSON.stringify(newoptions_division));
                    propPaneDivisions.unshift(blankOption);
                    // this.context.propertyPane.refresh();
                }
                else {
                    update(this.properties, 'prefilter_label_department', function () { return ''; });
                }
            }
        }
        if (this.properties.options_division) {
            if (this.properties.prefilter_key_division) {
                if (this.properties.prefilter_key_division != 'NoFilter') {
                    var newDivisionLabel_1 = this.properties.options_division.find(function (obj) { return obj.key == _this.properties.prefilter_key_division; }).text;
                    update(this.properties, 'prefilter_label_division', function () { return newDivisionLabel_1; });
                }
                else {
                    update(this.properties, 'prefilter_label_division', function () { return ''; });
                }
            }
        }
        // if (this.properties.show_department && this.properties.show_division && this.properties.show_department)
        var element = React.createElement("div", null,
            React.createElement(App, { context: this.context, searchBoxPlaceholder: this.properties.searchBoxPlaceholder, appHeading: this.properties.appHeading, initialResultText: this.properties.initialResultText, noResultText: this.properties.noResultText, show_department: this.properties.show_department, show_division: this.properties.show_division, show_organization: this.properties.show_organization, prefilter_key_department: this.properties.prefilter_key_department, prefilter_label_department: this.properties.prefilter_label_department, prefilter_key_division: this.properties.prefilter_key_division, prefilter_label_division: this.properties.prefilter_label_division, options_department: this.properties.options_department, options_division: this.properties.options_division, options_organization: availOrganizationsObject }));
        ReactDom.render(element, this.domElement);
    };
    PhoneListSearchWebPart.prototype.onDispose = function () {
        ReactDom.unmountComponentAtNode(this.domElement);
    };
    Object.defineProperty(PhoneListSearchWebPart.prototype, "dataVersion", {
        get: function () {
            return Version.parse('1.0');
        },
        enumerable: true,
        configurable: true
    });
    PhoneListSearchWebPart.prototype.getPropertyPaneConfiguration = function () {
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
                            groupName: 'Fields to Show in Results',
                            groupFields: [
                                PropertyPaneCheckbox('show_organization', {
                                    text: 'Organization'
                                }),
                                PropertyPaneCheckbox('show_department', {
                                    text: 'Department'
                                }),
                                PropertyPaneCheckbox('show_division', {
                                    text: 'Division'
                                })
                            ]
                        },
                        {
                            groupName: 'Preconfigured Filters',
                            groupFields: [
                                PropertyPaneDropdown('prefilter_key_department', {
                                    label: 'Departments',
                                    options: propPaneDepartments,
                                    selectedKey: this.properties.prefilter_key_department
                                }),
                                PropertyPaneDropdown('prefilter_key_division', {
                                    label: 'Divisions',
                                    options: propPaneDivisions,
                                    selectedKey: this.properties.prefilter_key_division
                                })
                            ]
                        }
                    ]
                }
            ]
        };
    };
    return PhoneListSearchWebPart;
}(BaseClientSideWebPart));
export default PhoneListSearchWebPart;
// export interface IAppProps {
//   appHeading: string;
//   searchBoxPlaceholder: string;
//   initialResultText: string;
//   noResultText: string;
//   show_organization: boolean;
//   show_department: boolean;
//   show_division: boolean;
//   prefilter_key_department: string;
//   prefilter_key_division: string;
//   prefilter_label_department: string;
//   prefilter_label_division: string;
//   options_department: any;
//   options_division: any;
// }
// export interface IAppState {
//   needUpdate: boolean;
//   items: any;
//   searchTerms?: string;
//   view?: string;
//   order?: string;
//   size?: string;
//   showPanel: boolean;
//   filters: string;
//   hasFiltersOrganization: boolean;
//   hasFiltersDepartment: boolean;
//   hasFiltersDivision: boolean;
//   clearFilters: boolean;
// }
// export interface IContactSearchBoxProps {
//   parentCallback;
//   view?: string;
//   order?: string;
//   size?: string;
//   showPanel: boolean;
//   filters: string;
//   hasFiltersOrganization: boolean;
//   hasFiltersDepartment: boolean;
//   hasFiltersDivision: boolean;
//   searchBoxPlaceholder: string;
//   show_organization: boolean;
//   show_department: boolean;
//   show_division: boolean;
//   clearFilters: boolean;
//   prefilter_key_department: string;
//   prefilter_key_division: string;
//   prefilter_label_department: string;
//   prefilter_label_division: string;
//   options_department: any;
//   options_division: any;
// }
// export interface IContactSearchBoxState {
//   searchTerms: any;
//   items: any;
//   view?: string;
//   order?: string;
//   size?: string;
//   needUpdate: boolean;
//   showPanel: boolean;
//   filters: string;
//   hasFiltersOrganization: boolean;
//   hasFiltersDepartment: boolean;
//   hasFiltersDivision: boolean;
//   show_organization: boolean;
//   show_department: boolean;
//   show_division: boolean;
//   clearFilters: boolean;
// }
// export interface IResult {
//   key: string;
//   FirstName: string;
//   Title: string;
//   JobTitle: string;
//   Organization: string;
//   Company: string;
//   Division: string;
//   Program: string;
//   Email: string;
//   WorkPhone: string;
//   WorkAddress: string;
// }
// export interface IContactCardGridProps {
//   items?: any;
//   searchTerms: string;
//   size?: string;
//   show_organization: boolean;
//   show_department: boolean;
//   show_division: boolean;
// }
// export interface IContactCardGridState {
//   items?: any;
//   show_organization: boolean;
//   show_department: boolean;
//   show_division: boolean;
//   size: string;
// }
// export interface IContactCardProps {
//   item?: any;
//   searchTerms: string;
//   size?: string;
//   show_organization: boolean;
//   show_department: boolean;
//   show_division: boolean;
// }
// export interface IContactCardState {
//   item?: any;
//   show_organization: boolean;
//   show_department: boolean;
//   show_division: boolean;
// }
// export interface IFacepileBasicExampleState {
//   numberOfFaces: any;
//   imagesFadeIn: boolean;
//   personaSize: PersonaSize;
// }
// export interface IFacepileBasicExampleProps {
//   personas: any;
//   personaSize: number;
// }
// export interface ICommandBarSearchControlsProps {
//   parentCallback;
//   view?: string;
//   order?: string;
//   size?: string;
//   showPanel: boolean;
//   filters: string;
//   clearFilters: boolean;
//   prefilter_key_department: string;
//   prefilter_key_division: string;
//   prefilter_label_department: string;
//   prefilter_label_division: string;
//   options_department: any;
//   options_division: any;
// }
// export interface ICommandBarSearchControlsState {
//   view?: string;
//   order?: string;
//   size?: string;
//   showPanel: boolean;
//   filters: string;
//   hasFiltersOrganization: boolean;
//   hasFiltersDepartment: boolean;
//   hasFiltersDivision: boolean;
//   clearFilters: boolean;
// }
// export interface IDetailsListCustomColumnsResultsProp {
//   parentCallback;
//   items?: IResult[];
//   searchTerms?: any;
//   order?: string;
//   show_organization: boolean;
//   show_department: boolean;
//   show_division: boolean;
// }
// export interface IDetailsListCustomColumnsResultsState {
//   sortedItems: any;
//   columns: IColumn[];
//   searchTerms?: any;
//   order?: string;
//   show_organization: boolean;
//   show_department: boolean;
//   show_division: boolean;
// }
// export interface IFilterPanelProps {
//   parentCallback?;
//   showPanel: boolean;
//   filters: string;
//   clearFilters: boolean;
//   prefilter_key_department: string;
//   prefilter_key_division: string;
//   prefilter_label_department: string;
//   prefilter_label_division: string;
//   options_department: any;
//   options_division: any;
// }
// export interface IFilterPanelState {
//   showPanel: boolean;
//   hasChoiceData: boolean;
//   filters: string;
//   filtersOrganization: any;
//   filtersDepartment: any;
//   filtersDivision: any;
//   clearFilters: boolean;
//   prefilter_key_department?: string;
//   prefilter_key_division?: string;
//   prefilter_label_department?: string;
//   prefilter_label_division?: string;
// }
// export class ContactCard extends React.Component<IContactCardProps, IContactCardState> {
//   constructor(props) {
//     super(props);
//     this.state = {
//       item: this.props.item,
//       show_organization: this.props.show_organization,
//       show_department: this.props.show_department,
//       show_division: this.props.show_division
//     };
//   }
//   public componentDidUpdate(previousProps: IContactCardProps, previousState: IContactCardState) {
//     if (previousState.item != this.props.item) {
//       this.setState({ item: this.props.item }, () => {
//       });
//     }
//     if (previousState.show_organization != this.props.show_organization) {
//       this.setState({ show_organization: this.props.show_organization }, () => {
//       });
//     }
//     if (previousState.show_department != this.props.show_department) {
//       this.setState({ show_department: this.props.show_department }, () => {
//       });
//     }
//     if (previousState.show_division != this.props.show_division) {
//       this.setState({ show_division: this.props.show_division }, () => {
//       });
//     }
//   }
//   public render() {
//     const searchTerms = this.props.searchTerms;
//     let highlightHits = (str) => {
//       for (let term of searchTerms) {
//         const searchTermRegex = new RegExp(term.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'), "ig");
//         const searchTermHighlighted = '<span style="background-color:yellow;">$&</span>';
//         str = str.replace(searchTermRegex, searchTermHighlighted);
//       }
//       return str;
//     };
//     return (
//       <div
//         key={this.props.item.Id}
//         className={this.props.size == 'large' ? styles.contactItem : [styles.contactItem, styles.small].join(' ')}
//         data-item-id={this.props.item.Id}
//       >
//         <div className={styles.contactItemImg}>
//           <Link href={"https://delve-gcc.office.com/?p=" + this.props.item.Email + "&v=work"} target="about:blank">
//             <PersonaCoin
//               text={this.props.item.FirstName != null ? this.props.item.FirstName + ' ' + this.props.item.Title : this.props.item.Title}
//               coinSize={this.props.size == 'large' ? 100 : 50}
//               showInitialsUntilImageLoads={true}
//             />
//           </Link>
//         </div>
//         <div className={styles.contactItemDetails}>
//           <div className={styles.padBottom}>
//             <Link href={"https://delve-gcc.office.com/?p=" + this.props.item.Email + "&v=work"} target="about:blank">
//               <div className={[styles.contactItemFullName, styles.contactItemFieldBody].join(' ')}
//                 dangerouslySetInnerHTML={{
//                   __html: highlightHits(this.props.item.FirstName != null ? this.props.item.FirstName + ' ' + this.props.item.Title : this.props.item.Title)
//                 }}
//               />
//             </Link>
//             {this.props.item.JobTitle != null
//               ? <div className={styles.contactItemFieldBody}
//                 dangerouslySetInnerHTML={{ __html: highlightHits(this.props.item.JobTitle) }} />
//               : ''
//             }
//           </div>
//           <div className={styles.padBottom}>
//             {this.props.item.Organization != null && this.state.show_organization
//               ? <div className={styles.contactItemFieldBody}
//                 dangerouslySetInnerHTML={{ __html: highlightHits(this.props.item.Organization) }} />
//               : ''
//             }
//             {this.props.item.Company != null && this.state.show_department
//               ? <div>
//                 <span className={styles.contactItemFieldLabel}>Department: </span>
//                 <span className={styles.contactItemFieldBody}
//                   dangerouslySetInnerHTML={{ __html: highlightHits(this.props.item.Company) }} />
//               </div>
//               : ''
//             }
//             {this.props.item.Division != null && this.state.show_division
//               ? <div>
//                 <span className={styles.contactItemFieldLabel}>Division: </span>
//                 <span className={styles.contactItemFieldBody}
//                   dangerouslySetInnerHTML={{ __html: highlightHits(this.props.item.Division) }} />
//               </div>
//               : ''}
//             {this.props.item.Program != null
//               ? <div>
//                 <span className={styles.contactItemFieldLabel}>Program: </span>
//                 <span className={styles.contactItemFieldBody}
//                   dangerouslySetInnerHTML={{ __html: highlightHits(this.props.item.Program) }} />
//               </div>
//               : ''}
//           </div>
//           {this.props.item.Email != null
//             ? <div className={styles.contactItemFieldBody}>
//               <a href={'mailto:' + this.props.item.Email}>
//                 {this.props.item.Email}
//               </a>
//             </div>
//             : ''
//           }
//           {this.props.item.WorkPhone != null
//             ? <div className={styles.contactItemFieldBody}>{this.props.item.WorkPhone}</div>
//             : ''
//           }
//           {this.props.item.WorkAddress != null
//             ? <div className={styles.contactItemFieldBody}>{this.props.item.WorkAddress}</div>
//             : ''
//           }
//         </div>
//       </div>
//     );
//   }
// }
// export class ContactCardGrid extends React.Component<IContactCardGridProps, IContactCardGridState> {
//   constructor(props) {
//     super(props);
//     this.state = {
//       items: this.props.items,
//       show_organization: this.props.show_organization,
//       show_department: this.props.show_department,
//       show_division: this.props.show_division,
//       size: this.props.size
//     };
//   }
//   public render() {
//     return (
//       <div>
//         {this.state.items.map((item) => {
//           return (
//             <ContactCard
//               item={item}
//               searchTerms={this.props.searchTerms}
//               size={this.props.size}
//               show_organization={this.props.show_organization}
//               show_department={this.props.show_department}
//               show_division={this.props.show_division}
//             />
//           );
//         })}
//       </div>
//     );
//   }
//   public componentDidUpdate(previousProps: IContactCardGridProps, previousState: IContactCardGridState) {
//     if (previousState.items != this.props.items) {
//       this.setState({ items: this.props.items });
//     }
//   }
//   private _getKey(item: any, index?: number): string {
//     return item.key;
//   }
// }
// export class DetailsListCustomColumnsResults extends React.Component<IDetailsListCustomColumnsResultsProp, IDetailsListCustomColumnsResultsState> {
//   constructor(props) {
//     super(props);
//     this.state = {
//       sortedItems: this.props.items,
//       columns: _buildColumns(this.props.items, this.props.show_organization, this.props.show_department, this.props.show_division),
//       searchTerms: this.props.searchTerms,
//       order: this.props.order,
//       show_organization: this.props.show_organization,
//       show_department: this.props.show_department,
//       show_division: this.props.show_division
//     };
//     this._renderItemColumn = this._renderItemColumn.bind(this);
//   }
//   public render() {
//     const { sortedItems, columns, searchTerms } = this.state;
//     columns.map(column => {
//       column.isResizable = true;
//       column.name = column.fieldName == 'Company' ? 'Department'
//         : column.fieldName == 'Title' ? 'Last Name'
//           : column.fieldName.replace(/([A-Z])/g, ' $1').trim();
//     });
//     return (
//       <ShimmeredDetailsList
//         items={sortedItems}
//         setKey="set"
//         columns={columns}
//         onRenderItemColumn={this._renderItemColumn}
//         onColumnHeaderClick={this._onColumnClick}
//         ariaLabelForSelectionColumn="Toggle selection"
//         ariaLabelForSelectAllCheckbox="Toggle selection for all items"
//         checkButtonAriaLabel="Row checkbox"
//         searchTerms={searchTerms}
//       />
//     );
//   }
//   public sendData = (order) => {
//     this.props.parentCallback(order);
//   }
//   public componentDidUpdate(previousProps: IDetailsListCustomColumnsResultsProp, previousState: IDetailsListCustomColumnsResultsState) {
//     if (previousState.sortedItems != this.props.items) {
//       this.setState({ sortedItems: this.props.items });
//     }
//     if (previousState.order != this.props.order) {
//       this.setState({ order: this.props.order });
//     }
//     if (previousState.show_organization != this.props.show_organization) {
//       this.setState({
//         show_organization: this.props.show_organization,
//         columns: _buildColumns(this.props.items, this.props.show_organization, this.props.show_department, this.props.show_division)
//       });
//     }
//     if (previousState.show_department != this.props.show_department) {
//       this.setState({
//         show_department: this.props.show_department,
//         columns: _buildColumns(this.props.items, this.props.show_organization, this.props.show_department, this.props.show_division)
//       });
//     }
//     if (previousState.show_division != this.props.show_division) {
//       this.setState({
//         show_division: this.props.show_division,
//         columns: _buildColumns(this.props.items, this.props.show_organization, this.props.show_department, this.props.show_division)
//       });
//     }
//   }
//   public _renderItemColumn(item: IResult, index: number, column: IColumn, searchTerms: any) {
//     const searchTermsToHighlight = this.props.searchTerms;
//     let highlightHits = (str) => {
//       for (let term of searchTermsToHighlight) {
//         const searchTermRegex = new RegExp(term.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'), "ig");
//         const searchTermHighlighted = '<span style="background-color:yellow;">$&</span>';
//         str = str.replace(searchTermRegex, searchTermHighlighted);
//       }
//       return str;
//     };
//     const fieldContent = item[column.fieldName as keyof IResult] as string;
//     switch (column.key) {
//       case 'FirstName': case 'Title':
//         return (
//           fieldContent != null
//             ? <Link href={"https://delve-gcc.office.com/?p=" + item.Email + "&v=work"} target="about:blank">
//               <span dangerouslySetInnerHTML={{ __html: highlightHits(fieldContent) }} />
//             </Link>
//             : ''
//         );
//       case 'JobTitle': case 'Organization': case 'Company': case 'Division': case 'Program':
//         return (
//           fieldContent != null
//             ? <span dangerouslySetInnerHTML={{ __html: highlightHits(fieldContent) }} />
//             : ''
//         );
//       default:
//         return <span>{fieldContent}</span>;
//     }
//   }
//   private _onColumnClick = (event: React.MouseEvent<HTMLElement>, column: IColumn): void => {
//     let isSortedDescending = column.isSortedDescending;
//     if (column.isSorted) {
//       isSortedDescending = !isSortedDescending;
//     }
//     this.setState({
//       order: column.fieldName
//     }, () => {
//       this.sendData(this.state.order);
//     });
//   }
// }
// export class FilterPanel extends React.Component<IFilterPanelProps, IFilterPanelState> {
//   constructor(props) {
//     super(props);
//     this.state = {
//       showPanel: this.props.showPanel,
//       hasChoiceData: false,
//       filters: '',
//       filtersOrganization: [],
//       filtersDepartment: [],
//       filtersDivision: [],
//       clearFilters: false,
//       prefilter_key_department: '',
//       prefilter_key_division: ''
//     };
//   }
//   public componentDidUpdate(previousProps: IFilterPanelProps, previousState: IFilterPanelState) {
//     if (previousState.showPanel != this.props.showPanel) {
//       this.setState({ showPanel: this.props.showPanel }, () => {
//         this.sendData(this.state.showPanel, this.state.filters, this.state.filtersOrganization.length, this.state.filtersDepartment.length, this.state.filtersDivision.length, this.state.clearFilters);
//       });
//     }
//     if (previousState.hasChoiceData === false && this.state.hasChoiceData === false) {
//       this.setState({ hasChoiceData: true });
//     }
//     if (previousState.clearFilters != this.props.clearFilters) {
//       this.setState({ clearFilters: this.props.clearFilters });
//     }
//     if (previousState.prefilter_key_department != this.props.prefilter_key_department) {
//       this.setState({
//         prefilter_key_department: this.props.prefilter_key_department,
//         prefilter_label_department: this.props.prefilter_label_department
//       },
//         this._applyFilters);
//     }
//     if (previousState.prefilter_key_division != this.props.prefilter_key_division) {
//       this.setState({
//         prefilter_key_division: this.props.prefilter_key_division,
//         prefilter_label_division: this.props.prefilter_label_division
//       },
//         this._applyFilters);
//     }
//   }
//   public sendData = (showPanel, filters, hasFiltersOrganization, hasFiltersDepartment, hasFiltersDivision, clearFilters) => {
//     this.props.parentCallback(showPanel, filters, hasFiltersOrganization, hasFiltersDepartment, hasFiltersDivision, clearFilters);
//   }
//   private _showPanel = (): void => {
//     this.setState({ showPanel: true });
//   }
//   private _hidePanel = (): void => {
//     this.setState(
//       { showPanel: false },
//       () => { this.sendData(this.state.showPanel, this.state.filters, this.state.filtersOrganization.length, this.state.filtersDepartment.length, this.state.filtersDivision.length, this.state.clearFilters); }
//     );
//   }
//   private _applyFilters = () => {
//     let restFilters = [];
//     let hasFiltersOrganization = false;
//     let hasFiltersDepartment = false;
//     let hasFiltersDivision = false;
//     if (this.state.prefilter_label_department) {
//       if (this.state.prefilter_label_department != 'NoFilter') {
//         const restFiltersDepartment = "Company eq '" + this.state.prefilter_label_department.split('&').join('%26') + "'";
//         restFilters.push(restFiltersDepartment);
//         hasFiltersDepartment = true;
//       }
//     }
//     else if (this.state.filtersDepartment.length) {
//       const restFiltersDepartment = "(Company eq '" + this.state.filtersDepartment.join("' or Company eq '") + "')";
//       restFilters.push(restFiltersDepartment);
//       hasFiltersDepartment = true;
//     }
//     if (this.state.prefilter_label_division) {
//       if (this.state.prefilter_label_division != 'NoFilter') {
//         const restFiltersDivision = "Division eq '" + this.state.prefilter_label_division.split('&').join('%26') + "'";
//         restFilters.push(restFiltersDivision);
//         hasFiltersDivision = true;
//       }
//     }
//     else if (this.state.filtersDivision.length) {
//       const restFiltersDivision = "(Division eq '" + this.state.filtersDivision.join("' or Division eq '") + "')";
//       restFilters.push(restFiltersDivision);
//       hasFiltersDivision = true;
//     }
//     if (this.state.filtersOrganization.length) {
//       const restFiltersOrganization = "(Organization eq '" + this.state.filtersOrganization.join("' or Organization eq '") + "')";
//       restFilters.push(restFiltersOrganization);
//       hasFiltersOrganization = true;
//     }
//     this.setState(
//       { filters: restFilters.join(' and ') },
//       () => {
//         this.sendData(this.state.showPanel, this.state.filters, hasFiltersOrganization, hasFiltersDepartment, hasFiltersDivision, this.state.clearFilters);
//       }
//     );
//   }
//   private _clearFilters = () => {
//     this.setState(
//       {
//         showPanel: false,
//         filters: '',
//         filtersOrganization: [],
//         filtersDepartment: [],
//         filtersDivision: [],
//         clearFilters: true
//       }, () => {
//         this.sendData(this.state.showPanel, this.state.filters, false, false, false, true);
//       }
//     );
//   }
//   private _onRenderFooterContent = () => {
//     const applyFilterIcon: IIconProps = { iconName: 'WaitlistConfirmMirrored' };
//     const hideFilterIcon: IIconProps = { iconName: 'Hide' };
//     const clearFilterIcon: IIconProps = { iconName: 'ClearFilter' };
//     return (
//       <div>
//         <DefaultButton
//           iconProps={applyFilterIcon}
//           onClick={this._applyFilters}
//         >Apply</DefaultButton>
//         <DefaultButton
//           iconProps={hideFilterIcon}
//           styles={{ root: { marginLeft: 15 } }}
//           onClick={this._hidePanel}
//         >Hide</DefaultButton>
//         <DefaultButton
//           iconProps={clearFilterIcon}
//           styles={{ root: { marginLeft: 15 } }}
//           onClick={this._clearFilters}
//         >Clear</DefaultButton>
//       </div>
//     );
//   }
//   private _onFilterChangeOrganization = (e) => {
//     if (e.target.checked) {
//       let newFilters = this.state.filtersOrganization;
//       newFilters.push(e.target.title.split('&').join('%26'));
//       this.setState({
//         filtersOrganization: newFilters
//       });
//     }
//   }
//   private _onFilterChangeDepartment = (e) => {
//     if (e.target.checked) {
//       let newFilters = this.state.filtersDepartment;
//       newFilters.push(e.target.title.split('&').join('%26'));
//       this.setState({
//         filtersDepartment: newFilters
//       });
//     }
//   }
//   private _onFilterChangeDivision = (e) => {
//     if (e.target.checked) {
//       let newFilters = this.state.filtersDivision;
//       newFilters.push(e.target.title.split('&').join('%26'));
//       this.setState({
//         filtersDivision: newFilters
//       });
//     }
//   }
//   public render() {
//     return (
//       <Panel
//         key={this.state.clearFilters ? 'ReRender' : 'noReRender'}
//         isOpen={this.state.showPanel}
//         closeButtonAriaLabel='Close'
//         isLightDismiss={true}
//         headerText='Filter Contacts'
//         onDismiss={this._hidePanel}
//         onRenderFooterContent={this._onRenderFooterContent}
//         isHiddenOnDismiss={true}
//         isFooterAtBottom={true}
//         type={PanelType.custom}
//         customWidth='420px'
//       >
//         <Dropdown
//           placeholder={
//             this.state.prefilter_key_department != null
//               && this.state.prefilter_key_department != undefined
//               && this.state.prefilter_key_department != 'NoFilter'
//               ? 'Filtered by ' + this.state.prefilter_label_department
//               : 'Select departments...'
//           }
//           label='Department'
//           multiSelect
//           options={this.props.options_department}
//           styles={{ dropdown: { width: 300 } }}
//           disabled={
//             this.state.prefilter_key_department != null
//             && this.state.prefilter_key_department != undefined
//             && this.state.prefilter_key_department != 'NoFilter'
//           }
//           onChange={this._onFilterChangeDepartment}
//         />
//         <Dropdown
//           placeholder={
//             this.state.prefilter_key_division != null
//               && this.state.prefilter_key_division != undefined
//               && this.state.prefilter_key_division != 'NoFilter'
//               ? 'Filtered by ' + this.state.prefilter_key_division
//               : 'Select divisions...'
//           }
//           label='Division'
//           multiSelect
//           options={this.props.options_division}
//           styles={{ dropdown: { width: 300 } }}
//           disabled={
//             this.state.prefilter_key_division != null &&
//             this.state.prefilter_key_division != undefined
//             && this.state.prefilter_key_division != 'NoFilter'
//           }
//           onChange={this._onFilterChangeDivision}
//         />
//         <Dropdown
//           placeholder='Select organizations...'
//           label='Organization'
//           multiSelect
//           options={availOrganizationsObject}
//           styles={{ dropdown: { width: 300 } }}
//           onChange={this._onFilterChangeOrganization}
//         />
//       </Panel>
//     );
//   }
// }
// export class CommandBarSearchControls extends React.Component<ICommandBarSearchControlsProps, ICommandBarSearchControlsState> {
//   constructor(props) {
//     super(props);
//     this.state = {
//       view: this.props.view,
//       order: this.props.order,
//       size: 'small',
//       showPanel: this.props.showPanel,
//       filters: this.props.filters,
//       hasFiltersOrganization: false,
//       hasFiltersDepartment: false,
//       hasFiltersDivision: false,
//       clearFilters: this.props.clearFilters
//     };
//     this.handleViewTilesClick = this.handleViewTilesClick.bind(this);
//     this.handleViewListClick = this.handleViewListClick.bind(this);
//     this.handleSortTilesClick = this.handleSortTilesClick.bind(this);
//     this.handleFilterClick = this.handleFilterClick.bind(this);
//   }
//   public componentDidUpdate(previousProps: ICommandBarSearchControlsProps, previousState: ICommandBarSearchControlsState) {
//     if (previousState.filters != this.props.filters) {
//       this.setState({ filters: this.props.filters }, () => {
//         this.sendData(true, this.state.view, this.state.order, this.state.size, this.state.showPanel, this.state.filters, this.state.hasFiltersOrganization, this.state.hasFiltersDepartment, this.state.hasFiltersDivision, false);
//       });
//     }
//     if (previousState.clearFilters != this.props.clearFilters) {
//       this.setState({
//         clearFilters: this.props.clearFilters,
//         showPanel: false
//       }, () => {
//         this.sendData(true, this.state.view, this.state.order, this.state.size, this.state.showPanel, this.state.filters, this.state.hasFiltersOrganization, this.state.hasFiltersDepartment, this.state.hasFiltersDivision, this.state.clearFilters);
//       });
//     }
//   }
//   public sendData = (boolVal, view, order, size, showPanel, filters, hasFiltersOrganization, hasFiltersDepartment, hasFiltersDivision, clearFilters) => {
//     this.props.parentCallback(boolVal, view, order, size, showPanel, filters, hasFiltersOrganization, hasFiltersDepartment, hasFiltersDivision, clearFilters);
//   }
//   public handleFilterClick = () => {
//     this.setState({
//       showPanel: !this.state.showPanel
//     });
//   }
//   public handleSortTilesClick = (orderClicked) => {
//     this.setState({
//       order: orderClicked
//     }, () => {
//       this.sendData(true, this.state.view, this.state.order, this.state.size, this.state.showPanel, this.state.filters, this.state.hasFiltersOrganization, this.state.hasFiltersDepartment, this.state.hasFiltersDivision, false);
//     });
//   }
//   public handleViewTilesClick = () => {
//     this.setState({
//       view: 'Tiles'
//     }, () => {
//       this.sendData(true, this.state.view, this.state.order, this.state.size, this.state.showPanel, this.state.filters, this.state.hasFiltersOrganization, this.state.hasFiltersDepartment, this.state.hasFiltersDivision, false);
//     });
//   }
//   public handleViewListClick = () => {
//     this.setState({
//       view: 'List'
//     }, () => {
//       this.sendData(true, this.state.view, this.state.order, this.state.size, this.state.showPanel, this.state.filters, this.state.hasFiltersOrganization, this.state.hasFiltersDepartment, this.state.hasFiltersDivision, false);
//     });
//   }
//   public handleTileSizeClick = (sizeClicked) => {
//     this.setState({
//       size: sizeClicked
//     }, () => {
//       this.sendData(true, this.state.view, this.state.order, this.state.size, this.state.showPanel, this.state.filters, this.state.hasFiltersOrganization, this.state.hasFiltersDepartment, this.state.hasFiltersDivision, false);
//     });
//   }
//   public callbackFromFilterPanelToCommandBar = (showPanel, filters, hasFiltersOrganization, hasFiltersDepartment, hasFiltersDivision, clearFilters) => {
//     this.setState({
//       showPanel: showPanel,
//       filters: filters,
//       hasFiltersOrganization: hasFiltersOrganization,
//       hasFiltersDepartment: hasFiltersDepartment,
//       hasFiltersDivision: hasFiltersDivision,
//       clearFilters: clearFilters
//     },
//       () => {
//         this.sendData(true, this.state.view, this.state.order, this.state.size, this.state.showPanel, this.state.filters, this.state.hasFiltersOrganization, this.state.hasFiltersDepartment, this.state.hasFiltersDivision, clearFilters);
//       }
//     );
//   }
//   private getItems = () => {
//     if (this.state.view == 'Tiles') {
//       return [
//         {
//           key: 'size',
//           name: 'Tile Size',
//           ariaLabel: 'Tile Size',
//           iconProps: {
//             iconName: 'SizeLegacy'
//           },
//           onClick: () => { this.handleViewListClick(); },
//           subMenuProps: {
//             items: [
//               {
//                 key: 'small',
//                 name: 'Small',
//                 iconProps: {
//                   iconName: 'GridViewSmall'
//                 },
//                 onClick: () => {
//                   this.handleTileSizeClick('small');
//                 }
//               },
//               {
//                 key: 'large',
//                 name: 'Large',
//                 iconProps: {
//                   iconName: 'GridViewMedium'
//                 },
//                 onClick: () => {
//                   this.handleTileSizeClick('large');
//                 }
//               }
//             ]
//           }
//         },
//         {
//           key: 'list',
//           name: 'Switch to List View',
//           ariaLabel: 'Switch to List View',
//           iconProps: {
//             iconName: 'GroupedList'
//           },
//           onClick: () => { this.handleViewListClick(); }
//         }
//       ];
//     }
//     return [
//       {
//         key: 'tile',
//         name: 'Switch to Grid View',
//         ariaLabel: 'Switch to Grid View',
//         iconProps: {
//           iconName: 'Tiles'
//         },
//         onClick: () => { this.handleViewTilesClick(); }
//       }
//     ];
//   }
//   private getFarItems = () => {
//     if (this.state.view == 'Tiles') {
//       return [
//         {
//           key: 'sort',
//           name: 'Sort',
//           ariaLabel: 'Sort',
//           iconProps: {
//             iconName: 'SortLines'
//           },
//           subMenuProps: {
//             items: [
//               {
//                 key: 'firstName',
//                 name: 'First Name',
//                 iconProps: {
//                   iconName: 'UserOptional'
//                 },
//                 onClick: () => {
//                   this.handleSortTilesClick('FirstName');
//                 }
//               },
//               {
//                 key: 'lastName',
//                 name: 'Last Name',
//                 iconProps: {
//                   iconName: 'UserOptional'
//                 },
//                 onClick: () => {
//                   this.handleSortTilesClick('Title');
//                 }
//               },
//               {
//                 key: 'organization',
//                 name: 'Organization',
//                 iconProps: {
//                   iconName: 'Org'
//                 },
//                 onClick: () => {
//                   this.handleSortTilesClick('Organization');
//                 }
//               },
//               {
//                 key: 'department',
//                 name: 'Department',
//                 iconProps: {
//                   iconName: 'Teamwork'
//                 },
//                 onClick: () => {
//                   this.handleSortTilesClick('Company');
//                 }
//               }
//             ]
//           }
//         },
//         {
//           key: 'filter',
//           name: 'Filter',
//           ariaLabel: 'Filter',
//           iconProps: {
//             iconName: 'Filter'
//           },
//           onClick: () => {
//             this.handleFilterClick();
//           }
//         }
//       ];
//     }
//     return [
//       {
//         key: 'filter',
//         name: 'Filter',
//         ariaLabel: 'Filter',
//         iconProps: {
//           iconName: 'Filter'
//         },
//         onClick: () => {
//           this.handleFilterClick();
//         }
//       }
//     ];
//   }
//   public render(): JSX.Element {
//     return (<div>
//       <CommandBar
//         items={this.getItems()}
//         farItems={this.getFarItems()}
//         ariaLabel={'Use left and right arrow keys to navigate between commands'}
//       />
//       <FilterPanel
//         parentCallback={this.callbackFromFilterPanelToCommandBar}
//         showPanel={this.state.showPanel}
//         filters={this.state.filters}
//         clearFilters={this.state.clearFilters}
//         prefilter_key_department={this.props.prefilter_key_department}
//         prefilter_key_division={this.props.prefilter_key_division}
//         prefilter_label_department={this.props.prefilter_label_department}
//         prefilter_label_division={this.props.prefilter_label_division}
//         options_department={this.props.options_department}
//         options_division={this.props.options_division}
//       />
//     </div>);
//   }
// }
// export class ContactSearchBox extends React.Component<IContactSearchBoxProps, IContactSearchBoxState> {
//   constructor(props) {
//     super(props);
//     this.state = {
//       searchTerms: [],
//       items: [],
//       view: this.props.view,
//       order: this.props.order,
//       needUpdate: false,
//       showPanel: false,
//       filters: this.props.filters,
//       hasFiltersOrganization: this.props.hasFiltersOrganization,
//       hasFiltersDepartment: this.props.hasFiltersDepartment,
//       hasFiltersDivision: this.props.hasFiltersDivision,
//       show_organization: this.props.show_organization,
//       show_department: this.props.show_department,
//       show_division: this.props.show_division,
//       clearFilters: this.props.clearFilters
//     };
//     this.handleChange = this.handleChange.bind(this);
//     this.handleClear = this.handleClear.bind(this);
//   }
//   public componentDidUpdate(previousProps: IContactSearchBoxProps, previousState: IContactSearchBoxState) {
//     if (previousState.order != this.props.order) {
//       this.setState({ order: this.props.order, needUpdate: true }, () => {
//         if (this.state.view == 'List') {
//           this.getRESTResults(this.state.searchTerms);
//         }
//       });
//     }
//     if (previousState.size != this.props.size) {
//       this.setState({ size: this.props.size, needUpdate: true }, () => {
//       });
//     }
//     if (previousState.showPanel != this.props.showPanel) {
//       this.setState({ showPanel: this.props.showPanel, needUpdate: true }, () => {
//       });
//     }
//     if (previousState.filters != this.state.filters) {
//       this.getRESTResults(this.state.searchTerms);
//     }
//     if (previousState.show_organization != this.props.show_organization) {
//       this.setState({ show_organization: this.props.show_organization, needUpdate: true }, () => {
//         this.getRESTResults(this.state.searchTerms);
//       });
//     }
//     if (previousState.show_department != this.props.show_department) {
//       this.setState({ show_department: this.props.show_department, needUpdate: true }, () => {
//         this.getRESTResults(this.state.searchTerms);
//       });
//     }
//     if (previousState.show_division != this.props.show_division) {
//       this.setState({ show_division: this.props.show_division, needUpdate: true }, () => {
//         this.getRESTResults(this.state.searchTerms);
//       });
//     }
//     if (previousState.clearFilters != this.props.clearFilters) {
//       this.setState({
//         clearFilters: this.props.clearFilters
//       });
//     }
//   }
//   public sendData = (boolVal, childData, searchTerms, view, order, size, showPanel, filters, hasFiltersOrganization, hasFiltersDepartment, hasFiltersDivision, clearFilters) => {
//     this.props.parentCallback(boolVal, childData, searchTerms, view, order, size, showPanel, filters, hasFiltersOrganization, hasFiltersDepartment, hasFiltersDivision, clearFilters);
//   }
//   public handleChange = debounce(e => {
//     if (e.length) {
//       this.getRESTResults(e);
//     }
//   }, 1000);
//   public getRESTResults(e) {
//     let searchTerms = [];
//     const myPromise = new Promise((resolve, reject) => {
//       if (e.constructor === Array) {
//         searchTerms = e;
//       }
//       else {
//         searchTerms = e.split(' ');
//       }
//       let searchFilters = [];
//       const searchFields = [
//         'Title',
//         'FirstName',
//         'JobTitle',
//         'Program'
//       ];
//       if (!this.state.hasFiltersOrganization && this.state.show_organization) {
//         searchFields.push('Organization');
//       }
//       if (!this.state.hasFiltersDepartment && this.state.show_department) {
//         searchFields.push('Company');
//       }
//       if (!this.state.hasFiltersDivision && this.state.show_division) {
//         searchFields.push('Division');
//       }
//       for (let term of searchTerms) {
//         let theseTerms = [];
//         for (let field of searchFields) {
//           theseTerms.push("substringof('" + term + "'," + field + ")");
//         }
//         searchFilters.push("(" + theseTerms.join(' or ') + ")");
//       }
//       const searchSourceUrl = "https://auroragov.sharepoint.com/sites/PhoneList";
//       const listName = "EmployeeContactList";
//       const select = "$select=Id,Title,FirstName,Email,Company,JobTitle,WorkPhone,WorkAddress,Division,Program,Organization,CellPhone";
//       const top = "$top=100";
//       const searchBarFilters = "(" + searchFilters.join(' and ') + ")";
//       const refiners = this.state.filters != null && this.state.filters.length ? this.state.filters + " and " : '';
//       const filter = "$filter=" + refiners + searchBarFilters;
//       const sortOrder = '$orderby=' + this.state.order;
//       const requestUrl = searchSourceUrl + "/_api/web/lists/getbytitle('" + listName + "')/items?" + select + "&" + top + "&" + filter + "&" + sortOrder;
//       appContext.spHttpClient.get(requestUrl, SPHttpClient.configurations.v1)
//         .then((response: SPHttpClientResponse) => {
//           if (response.ok) {
//             response.json().then((responseJSON) => {
//               if (responseJSON != null) {
//                 let items: any[] = responseJSON.value;
//                 resolve(items);
//               }
//               reject(new Error('Something went wrong.'));
//             });
//           }
//         });
//     });
//     const onResolved = (items) => {
//       this.setState({
//         items: items,
//         searchTerms: searchTerms,
//         view: this.props.view,
//         order: this.props.order,
//         size: this.props.size,
//         hasFiltersOrganization: this.props.hasFiltersOrganization,
//         hasFiltersDepartment: this.props.hasFiltersDepartment,
//         hasFiltersDivision: this.props.hasFiltersDivision
//       }, () => {
//         this.sendData(true, this.state.items, this.state.searchTerms, this.state.view, this.state.order, this.state.size, this.state.showPanel, this.state.filters, this.state.hasFiltersOrganization, this.state.hasFiltersDepartment, this.state.hasFiltersDivision, false);
//       });
//     };
//     const onRejected = (error) => console.log(error);
//     myPromise.then(onResolved, onRejected);
//   }
//   public handleClear(e) {
//     this.setState({
//       items: [],
//       searchTerms: '',
//       order: ''
//     },
//       () => {
//         this.sendData(true, this.state.items, this.state.searchTerms, this.state.view, this.state.order, this.state.size, this.state.showPanel, this.state.filters, this.state.hasFiltersOrganization, this.state.hasFiltersDepartment, this.state.hasFiltersDivision, false);
//       }
//     );
//   }
//   public callbackFromCommandBarToSearchBox = (boolVal, view, order, size, showPanel, filters, hasFiltersOrganization, hasFiltersDepartment, hasFiltersDivision, clearFilters) => {
//     this.setState({
//       view: view,
//       order: order,
//       needUpdate: boolVal,
//       size: size,
//       showPanel: showPanel,
//       filters: filters,
//       hasFiltersOrganization: hasFiltersOrganization,
//       hasFiltersDepartment: hasFiltersDepartment,
//       hasFiltersDivision: hasFiltersDivision,
//       clearFilters: clearFilters
//     },
//       () => {
//         this.sendData(true, this.state.items, this.state.searchTerms, this.state.view, this.state.order, this.state.size, this.state.showPanel, this.state.filters, this.state.hasFiltersOrganization, this.state.hasFiltersDepartment, this.state.hasFiltersDivision, clearFilters);
//         this.handleChange(this.state.searchTerms);
//       }
//     );
//   }
//   public render() {
//     const controls = this.state.items.length
//       ? <CommandBarSearchControls
//         parentCallback={this.callbackFromCommandBarToSearchBox}
//         view={this.state.view}
//         order={this.state.order}
//         showPanel={this.state.showPanel}
//         filters={this.state.filters}
//         clearFilters={this.state.clearFilters}
//         prefilter_key_department={this.props.prefilter_key_department}
//         prefilter_key_division={this.props.prefilter_key_division}
//         prefilter_label_department={this.props.prefilter_label_department}
//         prefilter_label_division={this.props.prefilter_label_division}
//         options_department={this.props.options_department}
//         options_division={this.props.options_division}
//       />
//       : '';
//     return (<div>
//       <SearchBox
//         underlined
//         placeholder={this.props.searchBoxPlaceholder}
//         onChange={this.handleChange}
//         onClear={this.handleClear}
//       />
//       {controls}
//     </div>);
//   }
// }
// export class App extends React.Component<IAppProps, IAppState> {
//   constructor(props) {
//     super(props);
//     this.state = {
//       needUpdate: false,
//       items: [],
//       searchTerms: '',
//       view: 'Tiles',
//       order: 'FirstName',
//       size: 'small',
//       showPanel: false,
//       filters: '',
//       hasFiltersOrganization: false,
//       hasFiltersDepartment: false,
//       hasFiltersDivision: false,
//       clearFilters: false
//     };
//     this.callbackFromSearchBoxToApp = this.callbackFromSearchBoxToApp.bind(this);
//   }
//   public callbackFromSearchBoxToApp = (boolVal, childData, searchTerms, view, order, size, showPanel, filters, hasFiltersOrganization, hasFiltersDepartment, hasFiltersDivision, clearFilters) => {
//     this.setState({
//       needUpdate: boolVal,
//       items: childData,
//       searchTerms: searchTerms,
//       view: view,
//       order: order,
//       size: size,
//       showPanel: showPanel,
//       filters: filters,
//       hasFiltersOrganization: hasFiltersOrganization,
//       hasFiltersDepartment: hasFiltersDepartment,
//       hasFiltersDivision: hasFiltersDivision,
//       clearFilters: clearFilters
//     });
//   }
//   public callbackFromDetailsListToApp = (order) => {
//     this.setState({
//       order: order
//     },
//       () => {
//       }
//     );
//   }
//   public render() {
//     const resultViewElement =
//       this.state.searchTerms.length ?
//         this.state.items.length ?
//           this.state.view == 'Tiles'
//             ? <ContactCardGrid
//               items={this.state.items}
//               searchTerms={this.state.searchTerms}
//               size={this.state.size}
//               show_organization={this.props.show_organization}
//               show_department={this.props.show_department}
//               show_division={this.props.show_division}
//             />
//             : <DetailsListCustomColumnsResults
//               parentCallback={this.callbackFromDetailsListToApp}
//               items={this.state.items}
//               searchTerms={this.state.searchTerms}
//               order={this.state.order}
//               show_organization={this.props.show_organization}
//               show_department={this.props.show_department}
//               show_division={this.props.show_division}
//             />
//           : <div>{this.props.noResultText}</div>
//         : <div>{this.props.initialResultText}</div>;
//     return (<div id="appRootWrap">
//       <h1>{this.props.appHeading}</h1>
//       <ContactSearchBox
//         parentCallback={this.callbackFromSearchBoxToApp}
//         view={this.state.view}
//         order={this.state.order}
//         size={this.state.size}
//         showPanel={this.state.showPanel}
//         filters={this.state.filters}
//         hasFiltersOrganization={this.state.hasFiltersOrganization}
//         hasFiltersDepartment={this.state.hasFiltersDepartment}
//         hasFiltersDivision={this.state.hasFiltersDivision}
//         searchBoxPlaceholder={this.props.searchBoxPlaceholder}
//         show_organization={this.props.show_organization}
//         show_department={this.props.show_department}
//         show_division={this.props.show_division}
//         clearFilters={this.state.clearFilters}
//         prefilter_key_department={this.props.prefilter_key_department}
//         prefilter_key_division={this.props.prefilter_key_division}
//         prefilter_label_department={this.props.prefilter_label_department}
//         prefilter_label_division={this.props.prefilter_label_division}
//         options_department={this.props.options_department}
//         options_division={this.props.options_division}
//       />
//       {resultViewElement}
//     </div>);
//   }
// }
// function _buildColumns(items: IResult[], show_organization, show_department, show_division): IColumn[] {
//   let theColumns = [];
//   items.map(item => {
//     theColumns.push({
//       FirstName: item.FirstName,
//       Title: item.Title,
//       JobTitle: item.JobTitle,
//       ...show_organization ? { Organization: item.Organization } : null,
//       ...show_department ? { Company: item.Company } : null,
//       ...show_division ? { Division: item.Division } : null,
//       Program: item.Program,
//       Email: item.Email,
//       WorkPhone: item.WorkPhone,
//       WorkAddress: item.WorkAddress
//     });
//   });
//   const columns = buildColumns(theColumns);
//   return columns;
// }
//# sourceMappingURL=PhoneListSearchWebPart.js.map
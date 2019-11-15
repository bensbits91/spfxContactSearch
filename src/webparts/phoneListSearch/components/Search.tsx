import * as React from 'react';
import { debounce } from 'lodash';
import { SearchBox } from 'office-ui-fabric-react/lib/SearchBox';


// let appContext;

export interface ISearchProps {
   handler: any;
   searchBoxPlaceholder: string;
   // view?: string;
   // order?: string;
   // size?: string;
   // showPanel: boolean;
   // filters: string;
   // hasFiltersOrganization: boolean;
   // hasFiltersDepartment: boolean;
   // hasFiltersDivision: boolean;
   // showOrganization: boolean;
   // show_department: boolean;
   // showDivision: boolean;
   // clearFilters: boolean;
   // prefilter_key_department: string;
   // prefilter_key_division: string;
   // prefilter_label_department: string;
   // prefilter_label_division: string;
   // departmentOptions: any;
   // divisionOptions: any;
 }
 
 export interface ISearchState {
   // searchTerms: any;
   // items: any;
   // view?: string;
   // order?: string;
   // size?: string;
   // needUpdate: boolean;
   // showPanel: boolean;
   // filters: string;
   // hasFiltersOrganization: boolean;
   // hasFiltersDepartment: boolean;
   // hasFiltersDivision: boolean;
   // showOrganization: boolean;
   // show_department: boolean;
   // showDivision: boolean;
   // clearFilters: boolean;
 }
 
 export default class Search extends React.Component<ISearchProps, ISearchState> {

   constructor(props) {
     super(props);
 
     this.state = {
      //  searchTerms: [],
      //  items: [],
      //  view: this.props.view,
      //  order: this.props.order,
      //  needUpdate: false,
      //  showPanel: false,
      //  filters: this.props.filters,
      //  hasFiltersOrganization: this.props.hasFiltersOrganization,
      //  hasFiltersDepartment: this.props.hasFiltersDepartment,
      //  hasFiltersDivision: this.props.hasFiltersDivision,
      //  showOrganization: this.props.showOrganization,
      //  show_department: this.props.show_department,
      //  showDivision: this.props.showDivision,
      //  clearFilters: this.props.clearFilters
     };
 
     this.handleChange = this.handleChange.bind(this);
     this.handleClear = this.handleClear.bind(this);
   }
 
   public componentDidUpdate(previousProps: ISearchProps, previousState: ISearchState) {
   //   if (previousState.order != this.props.order) {
   //     this.setState({ order: this.props.order, needUpdate: true }, () => {
   //       if (this.state.view == 'List') {
   //         this.getRESTResults(this.state.searchTerms);
   //       }
   //     });
   //   }
   //   if (previousState.size != this.props.size) {
   //     this.setState({ size: this.props.size, needUpdate: true }, () => {
   //     });
   //   }
   //   if (previousState.showPanel != this.props.showPanel) {
   //     this.setState({ showPanel: this.props.showPanel, needUpdate: true }, () => {
   //     });
   //   }
   //   if (previousState.filters != this.state.filters) {
   //     this.getRESTResults(this.state.searchTerms);
   //   }
   //   if (previousState.showOrganization != this.props.showOrganization) {
   //     this.setState({ showOrganization: this.props.showOrganization, needUpdate: true }, () => {
   //       this.getRESTResults(this.state.searchTerms);
   //     });
   //   }
   //   if (previousState.show_department != this.props.show_department) {
   //     this.setState({ show_department: this.props.show_department, needUpdate: true }, () => {
   //       this.getRESTResults(this.state.searchTerms);
   //     });
   //   }
   //   if (previousState.showDivision != this.props.showDivision) {
   //     this.setState({ showDivision: this.props.showDivision, needUpdate: true }, () => {
   //       this.getRESTResults(this.state.searchTerms);
   //     });
   //   }
   //   if (previousState.clearFilters != this.props.clearFilters) {
   //     this.setState({
   //       clearFilters: this.props.clearFilters
   //     });
   //   }
 
 
   }
 
   // public sendData = (/* boolVal, childData,  */searchTerms/* , view, order, size, showPanel, filters, hasFiltersOrganization, hasFiltersDepartment, hasFiltersDivision, clearFilters */) => {
   //   this.props.handler(/* boolVal, childData,  */searchTerms/* , view, order, size, showPanel, filters, hasFiltersOrganization, hasFiltersDepartment, hasFiltersDivision, clearFilters */);
   // }
 
   public handleChange = debounce(e => {
     if (e.length) {
      //  this.getRESTResults(e);
      this.props.handler(e);
     }
   }, 1000);
 
   // public getRESTResults(e) {
   //   let searchTerms = [];
   //   const myPromise = new Promise((resolve, reject) => {
   //     if (e.constructor === Array) {
   //       searchTerms = e;
   //     }
   //     else {
   //       searchTerms = e.split(' ');
   //     }
   //     let searchFilters = [];
   //     const searchFields = [
   //       'Title',
   //       'FirstName',
   //       'JobTitle',
   //       'Program'
   //     ];
   //     if (!this.state.hasFiltersOrganization && this.state.showOrganization) {
   //       searchFields.push('Organization');
   //     }
   //     if (!this.state.hasFiltersDepartment && this.state.show_department) {
   //       searchFields.push('Company');
   //     }
   //     if (!this.state.hasFiltersDivision && this.state.showDivision) {
   //       searchFields.push('Division');
   //     }
   //     for (let term of searchTerms) {
   //       let theseTerms = [];
   //       for (let field of searchFields) {
   //         theseTerms.push("substringof('" + term + "'," + field + ")");
   //       }
   //       searchFilters.push("(" + theseTerms.join(' or ') + ")");
   //     }
   //     const searchSourceUrl = "https://auroragov.sharepoint.com/sites/PhoneList";
   //     const listName = "EmployeeContactList";
   //     const select = "$select=Id,Title,FirstName,Email,Company,JobTitle,WorkPhone,WorkAddress,Division,Program,Organization,CellPhone";
   //     const top = "$top=100";
 
   //     const searchBarFilters = "(" + searchFilters.join(' and ') + ")";
 
   //     const refiners = this.state.filters != null && this.state.filters.length ? this.state.filters + " and " : '';
 
   //     const filter = "$filter=" + refiners + searchBarFilters;
   //     const sortOrder = '$orderby=' + this.state.order;
   //     const requestUrl = searchSourceUrl + "/_api/web/lists/getbytitle('" + listName + "')/items?" + select + "&" + top + "&" + filter + "&" + sortOrder;
   //    //  appContext.spHttpClient.get(requestUrl, SPHttpClient.configurations.v1)
   //    //    .then((response: SPHttpClientResponse) => {
   //    //      if (response.ok) {
   //    //        response.json().then((responseJSON) => {
   //    //          if (responseJSON != null) {
   //    //            let items: any[] = responseJSON.value;
   //    //            resolve(items);
   //    //          }
   //    //          reject(new Error('Something went wrong.'));
   //    //        });
   //    //      }
   //    //    });
   //   });
   //   const onResolved = (items) => {
 
   //     this.setState({
   //       items: items,
   //       searchTerms: searchTerms,
   //       view: this.props.view,
   //       order: this.props.order,
   //       size: this.props.size,
   //       hasFiltersOrganization: this.props.hasFiltersOrganization,
   //       hasFiltersDepartment: this.props.hasFiltersDepartment,
   //       hasFiltersDivision: this.props.hasFiltersDivision
   //     }, () => {
   //       this.sendData(true, this.state.items, this.state.searchTerms, this.state.view, this.state.order, this.state.size, this.state.showPanel, this.state.filters, this.state.hasFiltersOrganization, this.state.hasFiltersDepartment, this.state.hasFiltersDivision, false);
   //     });
   //   };
   //   const onRejected = (error) => console.log(error);
 
   //   myPromise.then(onResolved, onRejected);
   // }
 
   public handleClear(e) {
      this.props.handler(null);
   //   this.setState({
   //     items: [],
   //     searchTerms: '',
   //     order: ''
   //   },
   //     () => {
   //       this.sendData(true, this.state.items, this.state.searchTerms, this.state.view, this.state.order, this.state.size, this.state.showPanel, this.state.filters, this.state.hasFiltersOrganization, this.state.hasFiltersDepartment, this.state.hasFiltersDivision, false);
   //     }
   //   );
   }
 
   // public callbackFromCommandBarToSearchBox = (boolVal, view, order, size, showPanel, filters, hasFiltersOrganization, hasFiltersDepartment, hasFiltersDivision, clearFilters) => {
   //   this.setState({
   //     view: view,
   //     order: order,
   //     needUpdate: boolVal,
   //     size: size,
   //     showPanel: showPanel,
   //     filters: filters,
   //     hasFiltersOrganization: hasFiltersOrganization,
   //     hasFiltersDepartment: hasFiltersDepartment,
   //     hasFiltersDivision: hasFiltersDivision,
   //     clearFilters: clearFilters
   //   },
   //     () => {
   //       this.sendData(true, this.state.items, this.state.searchTerms, this.state.view, this.state.order, this.state.size, this.state.showPanel, this.state.filters, this.state.hasFiltersOrganization, this.state.hasFiltersDepartment, this.state.hasFiltersDivision, clearFilters);
   //       this.handleChange(this.state.searchTerms);
   //     }
   //   );
   // }
 
   public render() {
   //   const controls = /* this.state.items.length
   //     ? <CommandBarSearchControls
   //       handler={this.callbackFromCommandBarToSearchBox}
   //       view={this.state.view}
   //       order={this.state.order}
   //       showPanel={this.state.showPanel}
   //       filters={this.state.filters}
   //       clearFilters={this.state.clearFilters}
   //       prefilter_key_department={this.props.prefilter_key_department}
   //       prefilter_key_division={this.props.prefilter_key_division}
   //       prefilter_label_department={this.props.prefilter_label_department}
   //       prefilter_label_division={this.props.prefilter_label_division}
   //       departmentOptions={this.props.departmentOptions}
   //       divisionOptions={this.props.divisionOptions}
   //     />
   //     :  */'';
     return (<div>
       <SearchBox
         underlined
         placeholder={this.props.searchBoxPlaceholder}
         onChange={this.handleChange}
         onClear={this.handleClear}
       />
       {/* {controls} */}
     </div>);
   }
 
 }
 
  
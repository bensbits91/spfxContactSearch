import * as React from 'react';
import { Web } from '@pnp/sp';

import Search from './Search';
import Commands from './Commands';
import FilterPanel from './FilterPanel';
import Grid from './Grid';
import List from './List';

export interface IAppProps {
   context: any;
   appHeading: string;
   searchBoxPlaceholder: string;
   initialResultText: string;
   noResultText: string;

   show_department: boolean;
   show_division: boolean;
   show_organization: boolean;

   prefilter_key_department: string;
   prefilter_key_division: string;
   prefilter_label_department: string;
   prefilter_label_division: string;

   // filter_department: any;
   // filter_division: any;
   // filter_organization: any;

   options_department: any;
   options_division: any;
   options_organization: any;
}

export interface IAppState {
   needUpdate: boolean; // DON'T NEED THIS          RIGHT ????????????????????????????????????????????????????????????????????????????????????????????????????
   items?: any;
   searchTerms?: string;
   view?: string;
   order?: string;
   size?: string;
   showPanel: boolean;

   hasFilters_department: boolean; // DON'T NEED THIS          RIGHT ????????????????????????????????????????????????????????????????????????????????????????????????????
   hasFilters_organization: boolean; // DON'T NEED THIS          RIGHT ????????????????????????????????????????????????????????????????????????????????????????????????????
   hasFilters_division: boolean; // DON'T NEED THIS          RIGHT ????????????????????????????????????????????????????????????????????????????????????????????????????

   clearFilters: boolean; // DON'T NEED THIS          RIGHT ????????????????????????????????????????????????????????????????????????????????????????????????????

   // options_department: any;
   options_division: any;
   options_organization: any;

   filter_department?: any;
   filter_division?: any;
   filter_organization?: any;

   filter_search?: any;
   filter_panel?: string;
}


export default class App extends React.Component<IAppProps, IAppState> {

   constructor(props) {
      super(props);

      this.state = {
         needUpdate: false,
         // items: [],
         // searchTerms: '',
         view: 'Tiles',
         order: 'FirstName',
         size: 'small',
         showPanel: false,
         // filters: '',
         hasFilters_organization: false,
         hasFilters_department: false,
         hasFilters_division: false,
         clearFilters: false,

         options_division: this.props.options_division,
         options_organization: this.props.options_organization,

         // filter_panel
      };

      this.handler_searchBox = this.handler_searchBox.bind(this);
      this.handler_commands = this.handler_commands.bind(this);
      this.handler_filterPanel = this.handler_filterPanel.bind(this);
      this.handler_list = this.handler_list.bind(this);
   }

   public componentDidUpdate(previousProps: IAppProps, previousState: IAppState) {
      console.log('%c App -> componentDidUpdate -> previousProps', 'color:cyan', previousProps);
      console.log('%c App -> componentDidUpdate -> this.props', 'color:cyan', this.props);
      console.log('%c App -> componentDidUpdate -> previousState', 'color:cyan', previousState);
      console.log('%c App -> componentDidUpdate -> this.state', 'color:cyan', this.state);
      console.log('%c App -> componentDidUpdate -> this.state.filter_search', 'color:orange', this.state.filter_search);
      console.log('%c App -> componentDidUpdate -> this.state.filter_panel', 'color:orange', this.state.filter_panel);
      console.log('%c App -> componentDidUpdate -> this.state.prefilter_key_department', 'color:orange', this.props.prefilter_key_department);
      console.log('%c App -> componentDidUpdate -> this.state.prefilter_key_division', 'color:orange', this.props.prefilter_key_division);
   }

   public handler_searchBox = (e/* boolVal, childData,  searchTerms , view, order, size, showPanel, filters, hasFilters_organization, hasFilters_department, hasFilters_division, clearFilters */) => {
      const terms = e ? e.constructor === Array ? e : e.split(' ') : null;

      if (terms) {
         const searchFields = [
            'Title',
            'FirstName',
            'JobTitle',
            'Program'
         ];
         if (!this.state.hasFilters_department && this.props.show_department) {
            searchFields.push('Company');
         }
         if (!this.state.hasFilters_division && this.props.show_division) {
            searchFields.push('Division');
         }
         if (!this.state.hasFilters_organization && this.props.show_organization) {
            searchFields.push('Organization');
         }
         let filter_search_temp = [];
         // if (value) {
         // const terms = value;
         // for (let term of terms) {

         // async function buildFilter_search() {}

         let filter_search: string;

                                                /* 
                                                   COMMIT TO GIT HUB                                                                                     <===========================================
                                                   CLEAN UP CODE                                                                                         <===========================================
                                                   FIX APOSTROPHE ISSUE                                                                                  <===========================================
                                                */

         // let thePromise = 
         terms.map(term => {
            let theseTerms = [];
            for (let field of searchFields) {
               theseTerms.push("substringof('" + term/* .replace(/'/g,'') */ + "'," + field + ")"); //      THIS DOESN'T WORK -- HOW CAN I FIND O'KEEFE ????????????????????????????????????????????????????????????
            }
            filter_search_temp.push(/* "(" +  */theseTerms.join(' or ')/*  + ")" */);
            /* const  */filter_search = "(" + filter_search_temp.join(' and ') + ")";
            console.log('TCL: filter_search', filter_search);
            // }).then(filter_search => {
            //    this.setState({
            //       filter_search: filter_search,
            //       searchTerms: terms
            //    }, () => {
            //       this.getResults(/* 'terms', terms */);
            //    });
         });

         this.setState({
            filter_search: filter_search,
            searchTerms: terms
         }, () => {
            console.log('TCL: buildFilter_search -> this.state.filter_search', this.state.filter_search);
            console.log('TCL: buildFilter_search -> this.state.searchTerms', this.state.searchTerms);
            this.getResults(/* 'terms', terms */);
         });


      }
      else {
         this.setState({
            searchTerms: null
         });
      }


      // this.setState({
      // needUpdate: boolVal,
      // items: childData,
      // searchTerms: searchTerms,
      // view: view,
      // order: order,
      // size: size,
      // showPanel: showPanel,
      // filters: filters,
      // hasFilters_organization: hasFilters_organization,
      // hasFilters_department: hasFilters_department,
      // hasFilters_division: hasFilters_division,
      // clearFilters: clearFilters
      // });
   }

   public handler_commands = (event, value) => {
      console.log('%c App -> handler_commands -> event', 'color:orange', event);
      console.log('%c App -> handler_commands -> value', 'color:orange', value);

      if (event == 'size') {
         this.setState({
            size: value
         });
      }
      else if (event == 'view') {
         this.setState({
            view: value
         });
      }
      else if (event == 'filter') {
         this.setState({
            showPanel: value
            // showPanel: true
         });
      }
      else if (event == 'sort') {
         console.log('%c : handler_commands -> sort', 'color:chocolate');
         this.setState({
            order: value
         }, () => {
            this.getResults();
         });
      }
   }

   public handler_filterPanel = (event, value1, value2) => {
      console.log('%c App -> handler_filterPanel -> event', 'color:yellow', event);
      console.log('%c App -> handler_filterPanel -> value1', 'color:yellow', value1);
      console.log('%c App -> handler_filterPanel -> value2', 'color:yellow', value2);

      if (event == 'hide') {
         console.log('%c App -> handler_filterPanel -> hide', 'color: yellow');
         this.setState({
            showPanel: false
         });
      }

      else if (event == 'department') {
         console.log('%c handler_filterPanel -> this.props', 'background-color:black', this.props);
         console.log('%c handler_filterPanel -> this.state', 'background-color:black', this.state);

         let f = this.state.filter_department ? JSON.parse(JSON.stringify(this.state.filter_department)) : []; // currently selected options with spaces
         console.log('%c handler_filterPanel -> f', 'background-color:darkolivegreen', f);

         const d_props = JSON.parse(JSON.stringify(this.props.options_division)); // original division options
         let d_state = JSON.parse(JSON.stringify(this.state.options_division)); // currently available division options

         const o_props = JSON.parse(JSON.stringify(this.props.options_organization)); // original organization options
         let o_state = JSON.parse(JSON.stringify(this.state.options_organization)); // currently available organization options

         if (value2) { // if the clicked department is now selected
            f.push(value1); // add it to the filter WITH spaces ??????????????????????????
         }
         else { // if the clicked department is now NOT selected
            f = f.filter(option => option != value1); // only leave options that don't match the clicked department WITHOUT spaces ?????????????????????????
         }
         console.log('%c handler_filterPanel -> f', 'background-color:indigo', f);

         if (f.length) { // if there are any department filters

            let f_nospaces = JSON.parse(JSON.stringify(f));
            f_nospaces = f_nospaces.map(n => {
               return n.replace(/ /g, '');
            });
            console.log('%c : handler_filterPanel -> f_nospaces', 'color:yellow;background-color:black', f_nospaces);

            d_state = d_props.filter(option => f_nospaces.indexOf(option.department) > -1); // only options from the ORIGINAL division options that are in the array of department filters
            console.log('%c App -> handler_filterPanel -> d_state', 'background-color:darkviolet', d_state);
            o_state = o_props.filter(option => f_nospaces.indexOf(option.department) > -1); // only options from the ORIGINAL org options that are in the array of department filters
            console.log('%c App -> handler_filterPanel -> o_state', 'background-color:darkviolet', o_state);
         }
         else { // if there are NO department filters
            d_state = d_props; // reset division options
            o_state = o_props; // reset org options
            console.log('%c should show all divisions and organizations now', 'color:yellow');
         }

         this.setState({
            filter_department: f,
            filter_division: null,
            filter_organization: null,

            options_division: d_state,
            options_organization: o_state
         });
      }


      else if (event == 'division') {
         let f = this.state.filter_division ? JSON.parse(JSON.stringify(this.state.filter_division)) : [];
         const o_props = JSON.parse(JSON.stringify(this.props.options_organization));
         let o_state = JSON.parse(JSON.stringify(this.state.options_organization));

         if (value2) {
            f.push(value1);
         }
         else {
            f = f.filter(option => option != value1.replace(/ /g, ''));
         }
         console.log('TCL: App -> handler_filterPanel -> f', f);

         if (f.length) {
            let f_nospaces = JSON.parse(JSON.stringify(f));
            f_nospaces = f_nospaces.map(n => {
               return n.replace(/ /g, '');
            });
            console.log('%c : handler_filterPanel -> f_nospaces', 'color:yellow;background-color:black', f_nospaces);
            o_state = o_props.filter(option => f_nospaces.indexOf(option.division) > -1/*  && f.indexOf(option.department) > -1 */);
            console.log('TCL: App -> handler_filterPanel -> o_state', o_state);
         }
         else {
            const fd = this.state.filter_department ? JSON.parse(JSON.stringify(this.state.filter_department)) : null;
            const fd_nospaces = fd ? fd.map(n => {
               return n.replace(/ /g, '');
            })
               : null;
            console.log('%c : handler_filterPanel -> fd_nospaces', 'color:pink;background-color:black', fd_nospaces);
            o_state = fd_nospaces.length ? o_props.filter(option => fd_nospaces.indexOf(option.department) > -1) : o_props;
            console.log('%c should show more/all organizations now', 'color:yellow');
         }

         this.setState({
            filter_division: f,
            filter_organization: null,

            options_organization: o_state
         });
      }

      else if (event == 'organization') {
         let f = this.state.filter_organization ? JSON.parse(JSON.stringify(this.state.filter_organization)) : [];

         if (value2) {
            f.push(value1);
         }
         else {
            f = f.filter(option => option != value1.replace(/ /g, ''));
         }
         console.log('TCL: App -> handler_filterPanel -> f', f);

         this.setState({
            filter_organization: f
         });
      }

      else if (event == 'apply') {

         let restFilters = [];
         let hasFilters_organization = false;
         let hasFilters_department = false;
         let hasFilters_division = false;

         /* if (this.props.prefilter_label_department) {
            if (this.props.prefilter_label_department != 'NoFilter') {
               const restFiltersDepartment = "Company eq '" + this.props.prefilter_label_department.split('&').join('%26') + "'";
               restFilters.push(restFiltersDepartment);
               hasFilters_department = true;
            }
         }
         else  */if (this.state.filter_department/* .length */) {
            const restFilter_department = "(Company eq '" + this.state.filter_department.join("' or Company eq '") + "')";
            restFilters.push(restFilter_department);
            hasFilters_department = true;
         }

         /* if (this.props.prefilter_label_division) {
            if (this.props.prefilter_label_division != 'NoFilter') {
               const restFiltersDivision = "Division eq '" + this.props.prefilter_label_division.split('&').join('%26') + "'";
               restFilters.push(restFiltersDivision);
               hasFilters_division = true;
            }
         }
         else  */if (this.state.filter_division/* .length */) {
            const restFilter_division = "(Division eq '" + this.state.filter_division.join("' or Division eq '") + "')";
            restFilters.push(restFilter_division);
            hasFilters_division = true;
         }

         if (this.state.filter_organization/* .length */) {
            const restFilter_organization = "(Organization eq '" + this.state.filter_organization.join("' or Organization eq '") + "')";
            restFilters.push(restFilter_organization);
            hasFilters_organization = true;
         }

         this.setState({
            filter_panel: restFilters.join(' and '),
            hasFilters_department: hasFilters_department,
            hasFilters_division: hasFilters_division,
            hasFilters_organization: hasFilters_organization
         }, () => {
            this.getResults();
         });
      }

      else if (event == 'clear') {
         this.setState({
            filter_department: null,
            filter_division: null,
            filter_organization: null,
            filter_panel: null,

            options_division: this.props.options_division,
            options_organization: this.props.options_organization
         }, () => {
            this.getResults();
         });
      }
   }

   public handler_list = (value) => {
      console.log('%c  App -> handler_list -> value', 'color:yellow', value);
      this.setState({
         order: value
      }, () => {
         this.getResults();
      });

      // this.setState({
      //    order: order
      // },
      //    () => {
      //    }
      // );
   }


   public getResults(/* what, value */) {

      /*       const searchFields = [
               'Title',
               'FirstName',
               'JobTitle',
               'Program'
            ];
            if (!this.state.hasFilters_department && this.props.show_department) {
               searchFields.push('Company');
            }
            if (!this.state.hasFilters_division && this.props.show_division) {
               searchFields.push('Division');
            }
            if (!this.state.hasFilters_organization && this.props.show_organization) {
               searchFields.push('Organization');
            }
      
            if (what = 'terms') {
               let filter_search = [];
               if (value) {
                  const terms = value;
                  // for (let term of terms) {
                  terms.map(term => {
                     let theseTerms = [];
                     for (let field of searchFields) {
                        theseTerms.push("substringof('" + term + "'," + field + ")");
                     }
                     filter_search.push("(" + theseTerms.join(' or ') + ")");
                  }).then(filter_search => {
                     this.setState({
                        filter_search: filter_search
                     });
                  });
                  // }
               }
               else {
      
      
      
      
               }
            }
       */
      // const searchBarFilters = "(" + this.state.filter_search.join(' and ') + ")";


      if (this.state.filter_search.length) {
         const select = 'Id,Title,FirstName,Email,Company,JobTitle,WorkPhone,WorkAddress,Division,Program,Organization,CellPhone';
         const orderBy = this.state.order;
         const orderByAsc = true;
         console.log('%c : getResults -> orderBy', 'color:hotpink', orderBy);


         console.log('%c : getResults -> this.state.filter_search', 'color:hotpink', this.state.filter_search);

         let filter_pre_array = [];
         if (this.props.prefilter_key_department && this.props.prefilter_key_department != 'NoFilter') {
            filter_pre_array.push("Company eq '" + this.props.prefilter_label_department + "'");
         }
         if (this.props.prefilter_key_division && this.props.prefilter_key_division != 'NoFilter') {
            filter_pre_array.push("Division eq '" + this.props.prefilter_label_division + "'");
         }
         console.log('%c : getResults -> filter_pre_array', 'color:lime', filter_pre_array);
         const filter_pre_string = filter_pre_array.length > 1 ? filter_pre_array.join(' and ') : filter_pre_array[0] || null;
         console.log('%c : getResults -> filter_pre_string', 'color:lime', filter_pre_string);


         let filter_panel_array = [];
         if (this.state.filter_panel) {
            filter_panel_array.push(this.state.filter_panel);
         }
         console.log('%c : getResults -> filter_panel_array', 'color:lime', filter_panel_array);
         const filter_panel_string = filter_panel_array.length > 1 ? filter_panel_array.join(' and ') : filter_panel_array[0] || null;
         console.log('%c : getResults -> filter_panel_string', 'color:lime', filter_panel_string);


         let filter_array = [this.state.filter_search];
         if (filter_panel_string) { filter_array.push(filter_panel_string); }
         if (filter_pre_string) { filter_array.push(filter_pre_string); }

         const filter = filter_array.length > 1 ? filter_array.join(' and ') : this.state.filter_search;
         console.log('%c : getResults -> filter', 'color:aqua', filter);




         const theWeb = new Web(this.props.context.pageContext.web.absoluteUrl);
         const theList = theWeb.lists.getByTitle('EmployeeContactList');
         const theItems = theList.items.select(select).orderBy(orderBy, orderByAsc).filter(filter).top(500);

         theItems.get/* All */().then(items => {
            this.setState({
               items: items//,
               //searchTerms: value
            });
         });
      }
      // else {
      //    console.log('NEED TO HANDLE EMPTY SEARCH <----------------------------------------------------------------');
      // }




   }

   // public sortCards(value) {
   //    console.log('%c : sortCards -> value', 'color:chocolate', value);

   // }

   public render() {

      const el_search =
         <Search
            handler={this.handler_searchBox}
            searchBoxPlaceholder={this.props.searchBoxPlaceholder}
         // view={this.state.view}
         // order={this.state.order}
         // size={this.state.size}
         // showPanel={this.state.showPanel}
         // filters={this.state.filters}
         // hasFilters_organization={this.state.hasFilters_organization}
         // hasFilters_department={this.state.hasFilters_department}
         // hasFilters_division={this.state.hasFilters_division}
         // show_organization={this.props.show_organization}
         // show_department={this.props.show_department}
         // show_division={this.props.show_division}
         // clearFilters={this.state.clearFilters}
         // prefilter_key_department={this.props.prefilter_key_department}
         // prefilter_key_division={this.props.prefilter_key_division}
         // prefilter_label_department={this.props.prefilter_label_department}
         // prefilter_label_division={this.props.prefilter_label_division}
         // options_departments={this.props.options_departments}
         // options_divisions={this.props.options_divisions}
         />;


      const el_commands = this.state.searchTerms ?
         <Commands
            handler={this.handler_commands}
            view={this.state.view}
            order={this.state.order}
            showPanel={this.state.showPanel}
         // filters={this.state.filters}
         // clearFilters={this.state.clearFilters}
         // prefilter_key_department={this.props.prefilter_key_department}
         // prefilter_key_division={this.props.prefilter_key_division}
         // prefilter_label_department={this.props.prefilter_label_department}
         // prefilter_label_division={this.props.prefilter_label_division}
         // options_department={this.props.options_department}
         // options_division={this.props.options_division}
         />
         : '';
      /* this.state.items.length
      ? <Commands
        handler={this.callbackFromCommandBarToSearchBox}
        view={this.state.view}
        order={this.state.order}
        showPanel={this.state.showPanel}
        filters={this.state.filters}
        clearFilters={this.state.clearFilters}
        prefilter_key_department={this.props.prefilter_key_department}
        prefilter_key_division={this.props.prefilter_key_division}
        prefilter_label_department={this.props.prefilter_label_department}
        prefilter_label_division={this.props.prefilter_label_division}
        options_department={this.props.options_department}
        options_division={this.props.options_division}
      />
      :  '';*/

      const el_filterPanel =
         <FilterPanel
            handler={this.handler_filterPanel}
            showPanel={this.state.showPanel}
            // filters={this.state.filters}
            // clearFilters={this.state.clearFilters}

            prefilter_key_department={this.props.prefilter_key_department}
            prefilter_key_division={this.props.prefilter_key_division}
            prefilter_label_department={this.props.prefilter_label_department}
            prefilter_label_division={this.props.prefilter_label_division}

            filter_department={this.state.filter_department}
            filter_division={this.state.filter_division}
            filter_organization={this.state.filter_organization}

            options_department={this.props.options_department}
            options_division={this.state.options_division}
            options_organization={this.state.options_organization}
         />;

      const el_results =
         this.state.searchTerms/* .length */ ?
            //    this.state.items.length ?
            this.state.items ?
               this.state.view == 'Tiles'
                  ? <Grid
                     items={this.state.items}
                     searchTerms={this.state.searchTerms}
                     size={this.state.size}
                     show_organization={this.props.show_organization}
                     show_department={this.props.show_department}
                     show_division={this.props.show_division}
                  />
                  : <List
                     handler={this.handler_list}
                     items={this.state.items}
                     searchTerms={this.state.searchTerms}
                     // order={this.state.order}
                     show_organization={this.props.show_organization}
                     show_department={this.props.show_department}
                     show_division={this.props.show_division}
                  />
               : <div>{this.props.noResultText}</div>
            : <div>{this.props.initialResultText}</div>;

      return (<div id="appRootWrap">
         <h1>{this.props.appHeading}</h1>
         {el_search}
         {el_commands}
         {el_filterPanel}
         {el_results}
      </div>);
   }

}


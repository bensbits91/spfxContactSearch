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

   options_department: any;
   options_division: any;
   options_organization: any;
}

export interface IAppState {
   items?: any;
   searchTerms?: string;
   view?: string;
   order?: string;
   size?: string;
   showPanel: boolean;

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
         view: 'Tiles',
         order: 'FirstName',
         size: 'small',
         showPanel: false,
         options_division: this.props.options_division,
         options_organization: this.props.options_organization,
      };

      this.handler_searchBox = this.handler_searchBox.bind(this);
      this.handler_commands = this.handler_commands.bind(this);
      this.handler_filterPanel = this.handler_filterPanel.bind(this);
      this.handler_list = this.handler_list.bind(this);
   }

   // public componentDidUpdate(previousProps: IAppProps, previousState: IAppState) {
   //    console.log('%c App -> componentDidUpdate -> previousProps', 'color:cyan', previousProps);
   //    console.log('%c App -> componentDidUpdate -> this.props', 'color:cyan', this.props);
   //    console.log('%c App -> componentDidUpdate -> previousState', 'color:cyan', previousState);
   //    console.log('%c App -> componentDidUpdate -> this.state', 'color:cyan', this.state);
   //    console.log('%c App -> componentDidUpdate -> this.state.filter_search', 'color:orange', this.state.filter_search);
   //    console.log('%c App -> componentDidUpdate -> this.state.filter_panel', 'color:orange', this.state.filter_panel);
   //    console.log('%c App -> componentDidUpdate -> this.state.prefilter_key_department', 'color:orange', this.props.prefilter_key_department);
   //    console.log('%c App -> componentDidUpdate -> this.state.prefilter_key_division', 'color:orange', this.props.prefilter_key_division);
   // }

   public handler_searchBox = (e) => {
      const terms = e ? e.constructor === Array ? e : e.split(' ') : null;

      if (terms) {
         const searchFields = [
            'Title',
            'FirstName',
            'JobTitle',
            'Program'
         ];
         if (!this.state.filter_department && this.props.show_department) {
            searchFields.push('Company');
         }
         if (!this.state.filter_division && this.props.show_division) {
            searchFields.push('Division');
         }
         if (!this.state.filter_organization && this.props.show_organization) {
            searchFields.push('Organization');
         }
         let filter_search_temp = [];
         let filter_search: string;

         terms.map(term => {
            const term_uri = encodeURIComponent(term).replace(/'/g, '%27%27');
            let theseTerms = [];
            for (let field of searchFields) {
               theseTerms.push("substringof('" + term_uri + "'," + field + ")");
            }
            filter_search_temp.push(theseTerms.join(' or '));
            filter_search = "(" + filter_search_temp.join(' and ') + ")";
         });

         this.setState({
            filter_search: filter_search,
            searchTerms: terms
         }, () => {
            this.getResults();
         });


      }
      else {
         this.setState({
            searchTerms: null
         });
      }

   }

   public handler_commands = (event, value) => {

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
         });
      }
      else if (event == 'sort') {
         this.setState({
            order: value
         }, () => {
            this.getResults();
         });
      }
   }

   public handler_filterPanel = (event, value1, value2) => {

      if (event == 'hide') {
         this.setState({
            showPanel: false
         });
      }

      else if (event == 'department') {
         let f = this.state.filter_department ? JSON.parse(JSON.stringify(this.state.filter_department)) : []; // currently selected options with spaces

         const d_props = JSON.parse(JSON.stringify(this.props.options_division)); // original division options
         let d_state = JSON.parse(JSON.stringify(this.state.options_division)); // currently available division options

         const o_props = JSON.parse(JSON.stringify(this.props.options_organization)); // original organization options
         let o_state = JSON.parse(JSON.stringify(this.state.options_organization)); // currently available organization options

         if (value2) { // if the clicked department is now selected
            f.push(value1); // add it to the filter WITH spaces
         }
         else { // if the clicked department is now NOT selected
            f = f.filter(option => option != value1); // only leave options that don't match the clicked department WITHOUT spaces
         }

         if (f.length) { // if there are any department filters

            let f_nospaces = JSON.parse(JSON.stringify(f));
            f_nospaces = f_nospaces.map(n => {
               return n.replace(/ /g, '');
            });

            d_state = d_props.filter(option => f_nospaces.indexOf(option.department) > -1); // only options from the ORIGINAL division options that are in the array of department filters
            o_state = o_props.filter(option => f_nospaces.indexOf(option.department) > -1); // only options from the ORIGINAL org options that are in the array of department filters
         }
         else { // if there are NO department filters
            d_state = d_props; // reset division options
            o_state = o_props; // reset org options
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
            f = f.filter(option => option != value1);
         }

         if (f.length) {
            let f_nospaces = JSON.parse(JSON.stringify(f));
            f_nospaces = f_nospaces.map(n => {
               return n.replace(/ /g, '');
            });
            o_state = o_props.filter(option => f_nospaces.indexOf(option.division) > -1);
         }
         else {
            const fd = this.state.filter_department ? JSON.parse(JSON.stringify(this.state.filter_department)) : null;
            const fd_nospaces = fd ?
               fd.map(n => {
                  return n.replace(/ /g, '');
               })
               : null;
            o_state = fd_nospaces ? o_props.filter(option => fd_nospaces.indexOf(option.department) > -1) : o_props;
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

         this.setState({
            filter_organization: f
         });
      }

      else if (event == 'apply') {

         let restFilters = [];

         if (this.state.filter_department) {
            const restFilter_department = this.state.filter_department.length ? "(Company eq '" + this.state.filter_department.join("' or Company eq '") + "')" : null;
            if (restFilter_department) {
               restFilters.push(restFilter_department);
            }
            
         }

         if (this.state.filter_division) {
            const restFilter_division = this.state.filter_division.length ? "(Division eq '" + this.state.filter_division.join("' or Division eq '") + "')" : null;
            restFilters.push(restFilter_division);
         }

         if (this.state.filter_organization) {
            const restFilter_organization = this.state.filter_organization.length ? "(Organization eq '" + this.state.filter_organization.join("' or Organization eq '") + "')" : null;
            restFilters.push(restFilter_organization);
         }

         this.setState({
            filter_panel: restFilters.join(' and '),
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
      this.setState({
         order: value
      }, () => {
         this.getResults();
      });
   }

   public getResults() {

      if (this.state.filter_search.length) {
         const select = 'Id,Title,FirstName,Email,Company,JobTitle,WorkPhone,WorkAddress,Division,Program,Organization,CellPhone';
         const orderBy = this.state.order;
         const orderByAsc = true;



         let filter_pre_array = [];
         if (this.props.prefilter_key_department && this.props.prefilter_key_department != 'NoFilter') {
            filter_pre_array.push("Company eq '" + this.props.prefilter_label_department + "'");
         }
         if (this.props.prefilter_key_division && this.props.prefilter_key_division != 'NoFilter') {
            filter_pre_array.push("Division eq '" + this.props.prefilter_label_division + "'");
         }
         const filter_pre_string = filter_pre_array.length > 1 ? filter_pre_array.join(' and ') : filter_pre_array[0] || null;


         let filter_panel_array = [];
         if (this.state.filter_panel) {
            filter_panel_array.push(this.state.filter_panel);
         }
         const filter_panel_string = filter_panel_array.length > 1 ? filter_panel_array.join(' and ') : filter_panel_array[0] || null;

         let filter_array = [this.state.filter_search];
         if (filter_panel_string) { filter_array.push(filter_panel_string); }
         if (filter_pre_string) { filter_array.push(filter_pre_string); }

         const filter = filter_array.length > 1 ? filter_array.join(' and ') : this.state.filter_search;

         const theWeb = new Web(this.props.context.pageContext.web.absoluteUrl);
         const theList = theWeb.lists.getByTitle('EmployeeContactList');
         const theItems = theList.items.select(select).orderBy(orderBy, orderByAsc).filter(filter).top(500);

         theItems.get().then(items => {
            this.setState({
               items: items
            });
         });
      }

   }

   public render() {

      const el_search =
         <Search
            handler={this.handler_searchBox}
            searchBoxPlaceholder={this.props.searchBoxPlaceholder}
         />;


      const el_commands = this.state.searchTerms ?
         <Commands
            handler={this.handler_commands}
            view={this.state.view}
            order={this.state.order}
            showPanel={this.state.showPanel}
         />
         : '';

      const el_filterPanel =
         <FilterPanel
            handler={this.handler_filterPanel}
            showPanel={this.state.showPanel}

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
         this.state.searchTerms ?
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


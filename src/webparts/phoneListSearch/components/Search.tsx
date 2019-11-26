import * as React from 'react';
import { debounce } from 'lodash';
import { SearchBox } from 'office-ui-fabric-react/lib/SearchBox';


interface ISearchProps {
   handler: any;
   searchBoxPlaceholder: string;
 }
 
 interface ISearchState {

 }
 
 export default class Search extends React.Component<ISearchProps, ISearchState> {

   constructor(props) {
     super(props);
 
     this.state = {

     };
 
     this.handleChange = this.handleChange.bind(this);
     this.handleClear = this.handleClear.bind(this);
   }
 
   public handleChange = debounce(e => {
     if (e.length) {
      this.props.handler(e);
     }
   }, 1000);
 
   public handleClear(e) {
      this.props.handler(null);
   }
 
 
   public render() {
     return (<div>
       <SearchBox
         underlined
         placeholder={this.props.searchBoxPlaceholder}
         onChange={this.handleChange}
         onClear={this.handleClear}
       />
     </div>);
   }
 
 }
 
  
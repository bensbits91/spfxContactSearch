import * as React from 'react';

import Card from './Card';




interface IGridProps {
   items?: any;
   searchTerms: string;
   size?: string;
   show_department: boolean;
   show_division: boolean;
   show_organization: boolean;
}

interface IGridState {

}

export default class Grid extends React.Component<IGridProps, IGridState> {

   constructor(props) {
      super(props);
      this.state = {

      };
   }

   public render() {
      return (
         <div>
            {this.props.items.map(item => {
               return (
                  <Card
                     item={item}
                     searchTerms={this.props.searchTerms}
                     size={this.props.size}
                     show_department={this.props.show_department}
                     show_division={this.props.show_division}
                     show_organization={this.props.show_organization}
                  />
               );
            })}
         </div>
      );
   }
}


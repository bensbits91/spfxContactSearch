import * as React from 'react';

import Card from './Card';




export interface IGridProps {
   items?: any;
   searchTerms: string;
   size?: string;
   show_department: boolean;
   show_division: boolean;
   show_organization: boolean;
}

export interface IGridState {
   // items?: any;
   // show_department: boolean;
   // show_division: boolean;
   // show_organization: boolean;
   // size: string;
}

export default class Grid extends React.Component<IGridProps, IGridState> {

   constructor(props) {
      super(props);
      this.state = {
         //  items: this.props.items,
         //  show_organization: this.props.show_organization,
         //  show_department: this.props.show_department,
         //  show_division: this.props.show_division,
         //  size: this.props.size
      };
   }

   public componentDidMount() {
      console.log('%c : Grid -> componentDidMount -> this.props', 'color:yellow', this.props);
   }

   // public componentDidUpdate(previousProps: IGridProps, previousState: IGridState) {
   // console.log('%c Grid -> componentDidUpdate -> previousProps', 'color:magenta', previousProps);
   // console.log('%c Grid -> componentDidUpdate -> this.props', 'color:magenta', this.props);
   //   if (previousState.items != this.props.items) {
   //     this.setState({ items: this.props.items });
   //   }
   // }

   // private _getKey(item: any, index?: number): string {
   //    return item.key;
   // }

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


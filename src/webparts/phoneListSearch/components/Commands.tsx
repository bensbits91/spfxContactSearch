import * as React from 'react';
import { CommandBar } from 'office-ui-fabric-react/lib/CommandBar';


/* export  */interface ICommandsProps {
   handler;
   view?: string;
   order?: string;
   size?: string;
   showPanel: boolean;
}

/* export  */interface ICommandsState {

}

export default class Commands extends React.Component<ICommandsProps, ICommandsState> {

   constructor(props) {
      super(props);

      this.state = {

      };

      this.handleViewTilesClick = this.handleViewTilesClick.bind(this);
      this.handleViewListClick = this.handleViewListClick.bind(this);
      this.handleSortTilesClick = this.handleSortTilesClick.bind(this);
      this.handleFilterClick = this.handleFilterClick.bind(this);
   }

   public handleFilterClick = () => {
      this.props.handler('filter', true);
   }

   public handleSortTilesClick = (orderClicked) => {
      this.props.handler('sort', orderClicked);
   }

   public handleViewTilesClick = () => {
      this.props.handler('view', 'Tiles');
   }

   public handleViewListClick = () => {
      this.props.handler('view', 'List');
   }

   public handleTileSizeClick = (sizeClicked) => {
      this.props.handler('size', sizeClicked);
   }

   private getItems = () => {
      if (this.props.view == 'Tiles') {
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

   private getFarItems = () => {
      if (this.props.view == 'Tiles') {
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
      </div>);
   }

}


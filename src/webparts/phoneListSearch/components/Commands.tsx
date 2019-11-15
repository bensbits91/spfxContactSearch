import * as React from 'react';
import { CommandBar } from 'office-ui-fabric-react/lib/CommandBar';




export interface ICommandsProps {
   handler;
   view?: string;
   order?: string;
   size?: string;
   showPanel: boolean;
   // filters: string;
   // clearFilters: boolean;
   // prefilter_key_department: string;
   // prefilter_key_division: string;
   // prefilter_label_department: string;
   // prefilter_label_division: string;
   // departmentOptions: any;
   // divisionOptions: any;
}

export interface ICommandsState {
   // view?: string;
   // order?: string;
   // size?: string;
   // showPanel: boolean;
   // filters: string;
   // hasFiltersOrganization: boolean;
   // hasFiltersDepartment: boolean;
   // hasFiltersDivision: boolean;
   // clearFilters: boolean;
}

export default class Commands extends React.Component<ICommandsProps, ICommandsState> {

   constructor(props) {
      super(props);

      this.state = {
         //  view: this.props.view,
         //  order: this.props.order,
         //  size: 'small',
         //  showPanel: this.props.showPanel,
         //  filters: this.props.filters,
         //  hasFiltersOrganization: false,
         //  hasFiltersDepartment: false,
         //  hasFiltersDivision: false,
         //  clearFilters: this.props.clearFilters
      };

      this.handleViewTilesClick = this.handleViewTilesClick.bind(this);
      this.handleViewListClick = this.handleViewListClick.bind(this);
      this.handleSortTilesClick = this.handleSortTilesClick.bind(this);
      this.handleFilterClick = this.handleFilterClick.bind(this);
   }

   public componentDidUpdate(previousProps: ICommandsProps, previousState: ICommandsState) {
      //   if (previousState.filters != this.props.filters) {
      //     this.setState({ filters: this.props.filters }, () => {
      //       this.sendData(true, this.state.view, this.state.order, this.state.size, this.state.showPanel, this.state.filters, this.state.hasFiltersOrganization, this.state.hasFiltersDepartment, this.state.hasFiltersDivision, false);
      //     });
      //   }
      //   if (previousState.clearFilters != this.props.clearFilters) {
      //     this.setState({
      //       clearFilters: this.props.clearFilters,
      //       showPanel: false
      //     }, () => {
      //       this.sendData(true, this.state.view, this.state.order, this.state.size, this.state.showPanel, this.state.filters, this.state.hasFiltersOrganization, this.state.hasFiltersDepartment, this.state.hasFiltersDivision, this.state.clearFilters);
      //     });
      //   }
   }

   // public sendData = (boolVal, view, order, size, showPanel, filters, hasFiltersOrganization, hasFiltersDepartment, hasFiltersDivision, clearFilters) => {
   //   this.props.handler(boolVal, view, order, size, showPanel, filters, hasFiltersOrganization, hasFiltersDepartment, hasFiltersDivision, clearFilters);
   // }

   public handleFilterClick = () => {
      console.log('TCL: handleFilterClick -> handleFilterClick');
      this.props.handler('filter', true);
      //   this.setState({
      //     showPanel: !this.state.showPanel
      //   });
   }

   public handleSortTilesClick = (orderClicked) => {
      console.log('TCL: handleSortTilesClick -> orderClicked', orderClicked);
      this.props.handler('sort', orderClicked);
      //   this.setState({
      //     order: orderClicked
      //   }, () => {
      //     this.sendData(true, this.state.view, this.state.order, this.state.size, this.state.showPanel, this.state.filters, this.state.hasFiltersOrganization, this.state.hasFiltersDepartment, this.state.hasFiltersDivision, false);
      //   });
   }

   public handleViewTilesClick = () => {
      console.log('TCL: handleViewTilesClick -> handleViewTilesClick');
      this.props.handler('view', 'Tiles');
      // this.setState({
      //    view: 'Tiles'
      // }, () => {
      //    this.sendData(true, this.state.view, this.state.order, this.state.size, this.state.showPanel, this.state.filters, this.state.hasFiltersOrganization, this.state.hasFiltersDepartment, this.state.hasFiltersDivision, false);
      // });
   }

   public handleViewListClick = () => {
      console.log('TCL: handleViewListClick -> handleViewListClick');
      this.props.handler('view', 'List');
      // this.setState({
      //    view: 'List'
      // }, () => {
      //    this.sendData(true, this.state.view, this.state.order, this.state.size, this.state.showPanel, this.state.filters, this.state.hasFiltersOrganization, this.state.hasFiltersDepartment, this.state.hasFiltersDivision, false);
      // });
   }

   public handleTileSizeClick = (sizeClicked) => {
      console.log('TCL: handleTileSizeClick -> sizeClicked', sizeClicked);
      this.props.handler('size', sizeClicked);
      // this.setState({
      //    size: sizeClicked
      // }, () => {
      //    this.sendData(true, this.state.view, this.state.order, this.state.size, this.state.showPanel, this.state.filters, this.state.hasFiltersOrganization, this.state.hasFiltersDepartment, this.state.hasFiltersDivision, false);
      // });
   }

   // public callbackFromFilterPanelToCommandBar = (showPanel, filters, hasFiltersOrganization, hasFiltersDepartment, hasFiltersDivision, clearFilters) => {
   //    this.setState({
   //       showPanel: showPanel,
   //       filters: filters,
   //       hasFiltersOrganization: hasFiltersOrganization,
   //       hasFiltersDepartment: hasFiltersDepartment,
   //       hasFiltersDivision: hasFiltersDivision,
   //       clearFilters: clearFilters
   //    },
   //       () => {
   //          this.sendData(true, this.state.view, this.state.order, this.state.size, this.state.showPanel, this.state.filters, this.state.hasFiltersOrganization, this.state.hasFiltersDepartment, this.state.hasFiltersDivision, clearFilters);
   //       }
   //    );
   // }

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
         {/* <FilterPanel
         handler={this.callbackFromFilterPanelToCommandBar}
         showPanel={this.state.showPanel}
         filters={this.state.filters}
         clearFilters={this.state.clearFilters}
         prefilter_key_department={this.props.prefilter_key_department}
         prefilter_key_division={this.props.prefilter_key_division}
         prefilter_label_department={this.props.prefilter_label_department}
         prefilter_label_division={this.props.prefilter_label_division}
         departmentOptions={this.props.departmentOptions}
         divisionOptions={this.props.divisionOptions}
       /> */}
      </div>);
   }

}


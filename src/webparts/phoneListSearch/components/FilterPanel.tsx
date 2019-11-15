import * as React from 'react';
import { IIconProps } from 'office-ui-fabric-react/lib/Icon';
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { Dropdown } from 'office-ui-fabric-react/lib/Dropdown';



export interface IFilterPanelProps {
   handler: any;
   showPanel: boolean;
   // filters: string;
   // clearFilters: boolean;

   options_department: any;
   options_division: any;
   options_organization: any;

   prefilter_key_department: string;
   prefilter_key_division: string;
   prefilter_label_department: string;
   prefilter_label_division: string;

   filter_department?: any;
   filter_division?: any;
   filter_organization?: any;
}

export interface IFilterPanelState {
   // showPanel: boolean;
   // hasChoiceData: boolean;
   // filters: string;
   // filtersDepartment: any;
   // filtersDivision: any;
   // filtersOrganization: any;
   // clearFilters: boolean;
   // prefilter_key_department?: string;
   // prefilter_key_division?: string;
   // prefilter_label_department?: string;
   // prefilter_label_division?: string;
}



export default class FilterPanel extends React.Component<IFilterPanelProps, IFilterPanelState> {
   constructor(props) {
      super(props);

      this.state = {
         // showPanel: this.props.showPanel,
         // hasChoiceData: false,
         // filters: '',
         // filtersOrganization: [],
         // filtersDepartment: [],
         // filtersDivision: [],
         // clearFilters: false,
         // prefilter_key_department: '',
         // prefilter_key_division: ''
      };
   }




   /* 

   instead of setting state in this component,
   handler stuff back to App
   App function to get division/orgs based on selection
   App get new data
   AND re-render filter pane


   */





   public componentDidUpdate(previousProps: IFilterPanelProps, previousState: IFilterPanelState) {
      console.log('%c FilterPanel -> componentDidUpdate -> previousProps', 'color:magenta', previousProps);
      console.log('%c FilterPanel -> componentDidUpdate -> this.props', 'color:magenta', this.props);
      // if (previousState.showPanel != this.props.showPanel) {
      //    this.setState({ showPanel: this.props.showPanel }, () => {
      //       this.props.handler(this.state.showPanel, this.state.filters, this.state.filtersOrganization.length, this.state.filtersDepartment.length, this.state.filtersDivision.length, this.state.clearFilters);
      //    });
      // }

      // if (previousState.hasChoiceData === false && this.state.hasChoiceData === false) {
      //    this.setState({ hasChoiceData: true });
      // }

      // if (previousState.clearFilters != this.props.clearFilters) {
      //    this.setState({ clearFilters: this.props.clearFilters });
      // }

      // if (previousState.prefilter_key_department != this.props.prefilter_key_department) {
      //    this.setState({
      //       prefilter_key_department: this.props.prefilter_key_department,
      //       prefilter_label_department: this.props.prefilter_label_department
      //    },
      //       this._applyFilters);
      // }
      // if (previousState.prefilter_key_division != this.props.prefilter_key_division) {
      //    this.setState({
      //       prefilter_key_division: this.props.prefilter_key_division,
      //       prefilter_label_division: this.props.prefilter_label_division
      //    },
      //       this._applyFilters);
      // }
   }

   // public sendData = (showPanel, filters, hasFiltersOrganization, hasFiltersDepartment, hasFiltersDivision, clearFilters) => {
   //    this.props.handler(showPanel, filters, hasFiltersOrganization, hasFiltersDepartment, hasFiltersDivision, clearFilters);
   // }

   // private _showPanel = (): void => {
   //    this.setState({ showPanel: true });
   // }

   private _applyFilters = () => {
      this.props.handler('apply');

      // let restFilters = [];
      // let hasFiltersOrganization = false;
      // let hasFiltersDepartment = false;
      // let hasFiltersDivision = false;

      // if (this.state.prefilter_label_department) {
      //    if (this.state.prefilter_label_department != 'NoFilter') {
      //       const restFiltersDepartment = "Company eq '" + this.state.prefilter_label_department.split('&').join('%26') + "'";
      //       restFilters.push(restFiltersDepartment);
      //       hasFiltersDepartment = true;
      //    }
      // }
      // else if (this.state.filtersDepartment.length) {
      //    const restFiltersDepartment = "(Company eq '" + this.state.filtersDepartment.join("' or Company eq '") + "')";
      //    restFilters.push(restFiltersDepartment);
      //    hasFiltersDepartment = true;
      // }

      // if (this.state.prefilter_label_division) {
      //    if (this.state.prefilter_label_division != 'NoFilter') {
      //       const restFiltersDivision = "Division eq '" + this.state.prefilter_label_division.split('&').join('%26') + "'";
      //       restFilters.push(restFiltersDivision);
      //       hasFiltersDivision = true;
      //    }
      // }
      // else if (this.state.filtersDivision.length) {
      //    const restFiltersDivision = "(Division eq '" + this.state.filtersDivision.join("' or Division eq '") + "')";
      //    restFilters.push(restFiltersDivision);
      //    hasFiltersDivision = true;
      // }

      // if (this.state.filtersOrganization.length) {
      //    const restFiltersOrganization = "(Organization eq '" + this.state.filtersOrganization.join("' or Organization eq '") + "')";
      //    restFilters.push(restFiltersOrganization);
      //    hasFiltersOrganization = true;
      // }

      // this.setState(
      //    { filters: restFilters.join(' and ') },
      //    () => {
      //       this.props.handler(this.state.showPanel, this.state.filters, hasFiltersOrganization, hasFiltersDepartment, hasFiltersDivision, this.state.clearFilters);
      //    }
      // );
   }

   private _hidePanel = (): void => {
      this.props.handler('hide');
      // this.setState({
      //    showPanel: false
      // }, () => {
      // this.props.handler(this.state.showPanel, this.state.filters, this.state.filtersOrganization.length, this.state.filtersDepartment.length, this.state.filtersDivision.length, this.state.clearFilters);
      // });
   }

   private _clearFilters = () => {
      this.props.handler('clear');

      // this.setState(
      //    {
      //       showPanel: false,
      //       filters: '',
      //       filtersOrganization: [],
      //       filtersDepartment: [],
      //       filtersDivision: [],
      //       clearFilters: true
      //    }, () => {
      //       this.props.handler(this.state.showPanel, this.state.filters, false, false, false, true);
      //    }
      // );
   }

   private _onFilterChangeDepartment = (e) => {
      this.props.handler('department', e.target.title.split('&').join('%26'), e.target.checked);

      // this.setState({
      //    filtersDivision: []
      // });

      // if (e.target.checked) {
      //    let newFilters = this.state.filtersDepartment;
      //    newFilters.push(e.target.title.split('&').join('%26'));
      //    this.setState({
      //       filtersDepartment: newFilters
      //    });
      // }
   }

   private _onFilterChangeDivision = (e) => {
      this.props.handler('division', e.target.title.split('&').join('%26'), e.target.checked);

      // if (e.target.checked) {
      //    let newFilters = this.state.filtersDivision;
      //    newFilters.push(e.target.title.split('&').join('%26'));
      //    this.setState({
      //       filtersDivision: newFilters
      //    });
      // }
   }

   private _onFilterChangeOrganization = (e) => {
      this.props.handler('organization', e.target.title.split('&').join('%26'), e.target.checked);
      // if (e.target.checked) {
      //    let newFilters = this.state.filtersOrganization;
      //    newFilters.push(e.target.title.split('&').join('%26'));
      //    this.setState({
      //       filtersOrganization: newFilters
      //    });
      // }
   }

   private _onRenderFooterContent = () => {
      const applyFilterIcon: IIconProps = { iconName: 'WaitlistConfirmMirrored' };
      const hideFilterIcon: IIconProps = { iconName: 'Hide' };
      const clearFilterIcon: IIconProps = { iconName: 'ClearFilter' };
      return (
         <div>
            <DefaultButton
               iconProps={applyFilterIcon}
               // onClick={this.props.handler('apply')}
               onClick={this._applyFilters}
               text='Apply'
            />
            <DefaultButton
               iconProps={hideFilterIcon}
               styles={{ root: { marginLeft: 15 } }}
               // onClick={this.props.handler('hide')}
               onClick={this._hidePanel}
               text='Hide'
            />
            <DefaultButton
               iconProps={clearFilterIcon}
               styles={{ root: { marginLeft: 15 } }}
               // onClick={this.props.handler('clear')}
               onClick={this._clearFilters}
               text='Clear'
            />
         </div>
      );
   }

   public render() {
      let selectedKeys_department = [];
      if (this.props.filter_department) {
         this.props.filter_department.map(f => {
            selectedKeys_department.push(f.replace(/ /g, ''));
         });
         // selectedKeys_department = JSON.parse(JSON.stringify(this.props.filter_department));
         // selectedKeys_department.map(f => {
         //    /* return */ f.replace(/ /g, '');
         // });
      }

      let selectedKeys_division = [];
      if (this.props.filter_division) {
         this.props.filter_division.map(f => {
            selectedKeys_division.push(f.replace(/ /g, ''));
         });
         // selectedKeys_division = JSON.parse(JSON.stringify(this.props.filter_division));
         // selectedKeys_division.map(f => {
         //    /* return */ f.replace(/ /g, '');
         // });
      }

      let selectedKeys_organization = [];
      if (this.props.filter_organization) {
         this.props.filter_organization.map(f => {
            selectedKeys_organization.push(f.replace(/ /g, ''));
         });
         // selectedKeys_organization = JSON.parse(JSON.stringify(this.props.filter_organization));
         // selectedKeys_organization.map(f => {
         //    /* return */ f.replace(/ /g, '');
         // });
      }

      console.log('%c : render -> selectedKeys_department', 'background-color:darkred', selectedKeys_department);
      console.log('%c : render -> selectedKeys_division', 'background-color:darkred', selectedKeys_division);
      console.log('%c : render -> selectedKeys_organization', 'background-color:darkred', selectedKeys_organization);

      return (
         <Panel
            // key={this.state.clearFilters ? 'ReRender' : 'noReRender'}
            isOpen={this.props.showPanel}
            // isOpen={this.state.showPanel}
            closeButtonAriaLabel='Close'
            isLightDismiss={true}
            headerText='Filter Contacts'
            // onDismiss={this.props.handler('hide')}
            onDismiss={this._hidePanel}
            onRenderFooterContent={this._onRenderFooterContent}
            isHiddenOnDismiss={true}
            isFooterAtBottom={true}
            type={PanelType.custom}
            customWidth='420px'
         >
            <Dropdown
               placeholder={
                  this.props.prefilter_key_department
                     && this.props.prefilter_key_department != 'NoFilter'
                     ? 'Filtered by ' + this.props.prefilter_label_department
                     : 'Select departments...'
               }
               label='Department'
               multiSelect
               options={this.props.options_department}
               styles={{ dropdown: { width: 300 } }}
               disabled={
                  this.props.prefilter_key_department
                  && this.props.prefilter_key_department != 'NoFilter'
               }
               onChange={this._onFilterChangeDepartment}
               selectedKeys={selectedKeys_department}
            />
            <Dropdown
               placeholder={
                  this.props.prefilter_key_division
                     && this.props.prefilter_key_division != 'NoFilter'
                     ? 'Filtered by ' + this.props.prefilter_label_division
                     : 'Select divisions...'
               }
               label='Division'
               multiSelect
               options={this.props.options_division}
               styles={{ dropdown: { width: 300 } }}
               disabled={
                  this.props.prefilter_key_division
                  && this.props.prefilter_key_division != 'NoFilter'
               }
               onChange={this._onFilterChangeDivision}
               selectedKeys={selectedKeys_division}
            />
            <Dropdown
               placeholder='Select organizations...'
               label='Organization'
               multiSelect
               options={this.props.options_organization}
               styles={{ dropdown: { width: 300 } }}
               onChange={this._onFilterChangeOrganization}
               selectedKeys={selectedKeys_organization}
            />
         </Panel>
      );
   }

}


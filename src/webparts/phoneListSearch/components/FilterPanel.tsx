import * as React from 'react';
import { IIconProps } from 'office-ui-fabric-react/lib/Icon';
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { Dropdown } from 'office-ui-fabric-react/lib/Dropdown';



interface IFilterPanelProps {
   handler: any;
   showPanel: boolean;

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

interface IFilterPanelState {

}



export default class FilterPanel extends React.Component<IFilterPanelProps, IFilterPanelState> {
   constructor(props) {
      super(props);

      this.state = {

      };
   }

   private _applyFilters = () => {
      this.props.handler('apply');
   }

   private _hidePanel = (): void => {
      this.props.handler('hide');
   }

   private _clearFilters = () => {
      this.props.handler('clear');
   }

   private _onFilterChangeDepartment = (e) => {
      this.props.handler('department', e.target.title.split('&').join('%26'), e.target.checked);
   }

   private _onFilterChangeDivision = (e) => {
      this.props.handler('division', e.target.title.split('&').join('%26'), e.target.checked);
   }

   private _onFilterChangeOrganization = (e) => {
      this.props.handler('organization', e.target.title.split('&').join('%26'), e.target.checked);
   }

   private _onRenderFooterContent = () => {
      const applyFilterIcon: IIconProps = { iconName: 'WaitlistConfirmMirrored' };
      const hideFilterIcon: IIconProps = { iconName: 'Hide' };
      const clearFilterIcon: IIconProps = { iconName: 'ClearFilter' };
      return (
         <div>
            <DefaultButton
               iconProps={applyFilterIcon}
               onClick={this._applyFilters}
               text='Apply'
            />
            <DefaultButton
               iconProps={hideFilterIcon}
               styles={{ root: { marginLeft: 15 } }}
               onClick={this._hidePanel}
               text='Hide'
            />
            <DefaultButton
               iconProps={clearFilterIcon}
               styles={{ root: { marginLeft: 15 } }}
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
      }

      let selectedKeys_division = [];
      if (this.props.filter_division) {
         this.props.filter_division.map(f => {
            selectedKeys_division.push(f.replace(/ /g, ''));
         });
      }

      let selectedKeys_organization = [];
      if (this.props.filter_organization) {
         this.props.filter_organization.map(f => {
            selectedKeys_organization.push(f.replace(/ /g, ''));
         });
      }

      return (
         <Panel
            isOpen={this.props.showPanel}
            closeButtonAriaLabel='Close'
            isLightDismiss={true}
            headerText='Filter Contacts'
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


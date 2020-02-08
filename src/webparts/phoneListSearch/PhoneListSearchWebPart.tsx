

import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneCheckbox, PropertyPaneDropdown, IPropertyPaneCustomFieldProps, IPropertyPaneField, PropertyPaneFieldType } from '@microsoft/sp-property-pane';
import * as strings from 'PhoneListSearchWebPartStrings';
import PhoneListSearch from './components/PhoneListSearch';
import { IPhoneListSearchProps } from './components/IPhoneListSearchProps';
import styles from './components/PhoneListSearch.module.scss';
import { update } from '@microsoft/sp-lodash-subset';
import '@pnp/polyfill-ie11';
import { Web } from '@pnp/sp';
import App from './components/App';
import './components/temp.css';

// polyfills for IE11
import 'core-js/features/array/from';
import 'core-js/features/array/filter';
import 'regenerator-runtime/runtime';

export interface IPhoneListSearchWebPartProps {
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

   options_department: Array<any>;
   options_division: Array<any>;
   availOrganizationsObject: Array<any>;
}


let availOrganizationsObject = [];
let propPaneDepartments = [];
let propPaneDivisions = [];
const blankOption = {
   key: 'NoFilter',
   text: 'No Filter'
};


export default class PhoneListSearchWebPart extends BaseClientSideWebPart<IPhoneListSearchWebPartProps> {

   public availOrganizations = [];

   private getOptionsPromise: Promise<any>;

   public onInit(): Promise<void> {
      const theWeb = new Web(this.context.pageContext.web.absoluteUrl);
      const theList = theWeb.lists.getByTitle('EmployeeContactList');
      const select = 'Company,Division,Organization';

      this.getOptionsPromise = theList.items.select(select).getAll().then(items => {
         console.clear();
         console.log('TCL: PhoneListSearchWebPart -> items', items);

         let departmentsTempArray = [];
         let divisionsTempArray = [];

         update(this.properties, 'options_department', (): any => {
            return [];
         });
         update(this.properties, 'options_division', (): any => {
            return [];
         });
         update(this.properties, 'organizationOptions', (): any => {
            return [];
         });

         items.map(item => {
            if (item.Company != null) {
               if (departmentsTempArray.indexOf(item.Company) === -1) {
                  departmentsTempArray.push(item.Company);
                  this.properties.options_department.push({
                     key: item.Company.replace(/ /g, ''),
                     text: item.Company
                  });
               }
            }
            if (item.Division != null) {
               if (divisionsTempArray.indexOf(item.Division) === -1) {
                  divisionsTempArray.push(item.Division);
                  this.properties.options_division.push({
                     key: item.Division.replace(/ /g, ''),
                     text: item.Division,
                     department: item.Company.replace(/ /g, '')
                  });
               }
            }
            if (item.Organization != null) {
               if (this.availOrganizations.indexOf(item.Organization) === -1) {
                  this.availOrganizations.push(item.Organization);
                  availOrganizationsObject.push({
                     key: item.Organization.replace(/ /g, ''),
                     text: item.Organization,
                     department: item.Company.replace(/ /g, ''),
                     division: item.Division.replace(/ /g, '')
                  });
               }
            }
         });

         this.properties.options_department.sort(this.sortDropdowns);
         this.properties.options_division.sort(this.sortDropdowns);
         availOrganizationsObject.sort(this.sortDropdowns);
         
         propPaneDepartments = JSON.parse(JSON.stringify(this.properties.options_department));
         propPaneDepartments.unshift(blankOption);
         propPaneDivisions = JSON.parse(JSON.stringify(this.properties.options_division));
         propPaneDivisions.unshift(blankOption);
         
      // }).then(() => {
      //    console.clear();
      //    console.log('%c PhoneListSearchWebPart -> this.properties', 'color:aqua', this.properties);
      });
      
      return this.getOptionsPromise;
   }

   public sortDropdowns(a, b) {
      return (a.key > b.key) ? 1 : -1;
   }

   public render(): void {


      let newoptions_division = JSON.parse(JSON.stringify(this.properties.options_division));

      let newoptions_organization = JSON.parse(JSON.stringify(availOrganizationsObject));

      if (this.properties.options_department) {
         if (this.properties.prefilter_key_department) {
            if (this.properties.prefilter_key_department != 'NoFilter') {

               const newDeparmentLabel = this.properties.options_department.find(obj => obj.key == this.properties.prefilter_key_department).text;
               update(this.properties, 'prefilter_label_department', (): any => { return newDeparmentLabel; });

               newoptions_division = newoptions_division.filter(option => option.department == this.properties.prefilter_key_department);
               propPaneDivisions = JSON.parse(JSON.stringify(newoptions_division));
               propPaneDivisions.unshift(blankOption);

               newoptions_organization = newoptions_organization.filter(option => option.department == this.properties.prefilter_key_department);

            }
            else {
               update(this.properties, 'prefilter_label_department', (): any => { return ''; });
            }
         }
      }

      if (this.properties.options_division) {
         if (this.properties.prefilter_key_division) {
            if (this.properties.prefilter_key_division != 'NoFilter') {
               const newDivisionLabel = this.properties.options_division.find(obj => obj.key == this.properties.prefilter_key_division).text;
               update(this.properties, 'prefilter_label_division', (): any => { return newDivisionLabel; });

               newoptions_organization = newoptions_organization.filter(option => option.division == this.properties.prefilter_key_division);

            }
            else {
               update(this.properties, 'prefilter_label_division', (): any => { return ''; });
            }
         }
      }

      const element = <div>
         <App
            context={this.context}
            searchBoxPlaceholder={this.properties.searchBoxPlaceholder}
            appHeading={this.properties.appHeading}
            initialResultText={this.properties.initialResultText}
            noResultText={this.properties.noResultText}

            show_department={this.properties.show_department}
            show_division={this.properties.show_division}
            show_organization={this.properties.show_organization}

            prefilter_key_department={this.properties.prefilter_key_department}
            prefilter_label_department={this.properties.prefilter_label_department}
            prefilter_key_division={this.properties.prefilter_key_division}
            prefilter_label_division={this.properties.prefilter_label_division}

            options_department={this.properties.options_department}
            options_division={newoptions_division}
            options_organization={newoptions_organization}
         />
      </div>;

      ReactDom.render(element, this.domElement);
   }


   protected onDispose(): void {
      ReactDom.unmountComponentAtNode(this.domElement);
   }

   protected get dataVersion(): Version {
      return Version.parse('1.0');
   }

   protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {

      return {
         pages: [
            {
               header: {
                  description: strings.PropertyPaneDescription
               },
               groups: [
                  {
                     groupName: 'Page Text',
                     groupFields: [
                        PropertyPaneTextField('appHeading', {
                           label: 'Heading',
                           description: 'The heading that shows above the search box.'
                        }),
                        PropertyPaneTextField('searchBoxPlaceholder', {
                           label: 'Search Box Placeholder Text',
                           description: 'Text that shows inside the search box before the user enters text.'
                        }),
                        PropertyPaneTextField('initialResultText', {
                           label: 'Initial Result Text',
                           description: 'Text that shows in the results pane before the user searches.',
                           multiline: true
                        }),
                        PropertyPaneTextField('noResultText', {
                           label: 'No Result Text',
                           description: 'Text that shows in the results pane if no results are found.',
                           multiline: true
                        })
                     ]
                  },
                  {
                     groupName: 'Fields to Show in Results',
                     groupFields: [
                        PropertyPaneCheckbox('show_organization', {
                           text: 'Organization'
                        }),
                        PropertyPaneCheckbox('show_department', {
                           text: 'Department'
                        }),
                        PropertyPaneCheckbox('show_division', {
                           text: 'Division'
                        })
                     ]
                  },
                  {
                     groupName: 'Preconfigured Filters',
                     groupFields: [
                        PropertyPaneDropdown('prefilter_key_department', {
                           label: 'Departments',
                           options: propPaneDepartments,
                           selectedKey: this.properties.prefilter_key_department
                        }),
                        PropertyPaneDropdown('prefilter_key_division', {
                           label: 'Divisions',
                           options: propPaneDivisions,
                           selectedKey: this.properties.prefilter_key_division
                        })
                     ]
                  }
               ]
            }
         ]
      };
   }
}
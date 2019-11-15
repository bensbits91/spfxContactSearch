import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import './components/temp.css';
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
export default class PhoneListSearchWebPart extends BaseClientSideWebPart<IPhoneListSearchWebPartProps> {
    availOrganizations: any[];
    private getOptionsPromise;
    onInit(): Promise<void>;
    sortDropdowns(a: any, b: any): 1 | -1;
    render(): void;
    protected onDispose(): void;
    protected readonly dataVersion: Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
//# sourceMappingURL=PhoneListSearchWebPart.d.ts.map
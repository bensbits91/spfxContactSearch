/// <reference path="../../../node_modules/@types/lodash/common/common.d.ts" />
/// <reference path="../../../node_modules/@types/lodash/common/array.d.ts" />
/// <reference path="../../../node_modules/@types/lodash/common/collection.d.ts" />
/// <reference path="../../../node_modules/@types/lodash/common/date.d.ts" />
/// <reference path="../../../node_modules/@types/lodash/common/function.d.ts" />
/// <reference path="../../../node_modules/@types/lodash/common/lang.d.ts" />
/// <reference path="../../../node_modules/@types/lodash/common/math.d.ts" />
/// <reference path="../../../node_modules/@types/lodash/common/number.d.ts" />
/// <reference path="../../../node_modules/@types/lodash/common/object.d.ts" />
/// <reference path="../../../node_modules/@types/lodash/common/seq.d.ts" />
/// <reference path="../../../node_modules/@types/lodash/common/string.d.ts" />
/// <reference path="../../../node_modules/@types/lodash/common/util.d.ts" />
import * as React from 'react';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { IColumn } from 'office-ui-fabric-react/lib/DetailsList';
import { PersonaSize } from 'office-ui-fabric-react/lib/Persona';
export interface IPhoneListSearchWebPartProps {
    appHeading: string;
    searchBoxPlaceholder: string;
    initialResultText: string;
    noResultText: string;
    showOrganization: boolean;
    showDepartment: boolean;
    showDivision: boolean;
}
export interface IMainAppProps {
    appHeading: string;
    searchBoxPlaceholder: string;
    initialResultText: string;
    noResultText: string;
    showOrganization: boolean;
    showDepartment: boolean;
    showDivision: boolean;
}
export interface IMainAppState {
    needUpdate: boolean;
    items: any;
    searchTerms?: string;
    view?: string;
    order?: string;
    size?: string;
    showPanel: boolean;
    filters: string;
    hasFiltersOrganization: boolean;
    hasFiltersDepartment: boolean;
    hasFiltersDivision: boolean;
    clearFilters: boolean;
}
export interface IContactSearchBoxProps {
    parentCallback: any;
    view?: string;
    order?: string;
    size?: string;
    showPanel: boolean;
    filters: string;
    hasFiltersOrganization: boolean;
    hasFiltersDepartment: boolean;
    hasFiltersDivision: boolean;
    searchBoxPlaceholder: string;
    showOrganization: boolean;
    showDepartment: boolean;
    showDivision: boolean;
    clearFilters: boolean;
}
export interface IContactSearchBoxState {
    searchTerms: any;
    items: any;
    view?: string;
    order?: string;
    size?: string;
    needUpdate: boolean;
    showPanel: boolean;
    filters: string;
    hasFiltersOrganization: boolean;
    hasFiltersDepartment: boolean;
    hasFiltersDivision: boolean;
    showOrganization: boolean;
    showDepartment: boolean;
    showDivision: boolean;
    clearFilters: boolean;
}
export interface IResult {
    key: string;
    FirstName: string;
    Title: string;
    JobTitle: string;
    Organization: string;
    Company: string;
    Division: string;
    Program: string;
    Email: string;
    WorkPhone: string;
    WorkAddress: string;
}
export interface IContactCardGridProps {
    items?: any;
    searchTerms: string;
    size?: string;
    showOrganization: boolean;
    showDepartment: boolean;
    showDivision: boolean;
}
export interface IContactCardGridState {
    items?: any;
    showOrganization: boolean;
    showDepartment: boolean;
    showDivision: boolean;
}
export interface IContactCardProps {
    item?: any;
    searchTerms: string;
    size?: string;
    showOrganization: boolean;
    showDepartment: boolean;
    showDivision: boolean;
}
export interface IContactCardState {
    item?: any;
    showOrganization: boolean;
    showDepartment: boolean;
    showDivision: boolean;
}
export interface IFacepileBasicExampleState {
    numberOfFaces: any;
    imagesFadeIn: boolean;
    personaSize: PersonaSize;
}
export interface IFacepileBasicExampleProps {
    personas: any;
    personaSize: number;
}
export interface ICommandBarSearchControlsProps {
    parentCallback: any;
    view?: string;
    order?: string;
    size?: string;
    showPanel: boolean;
    filters: string;
    clearFilters: boolean;
}
export interface ICommandBarSearchControlsState {
    view?: string;
    order?: string;
    size?: string;
    showPanel: boolean;
    filters: string;
    hasFiltersOrganization: boolean;
    hasFiltersDepartment: boolean;
    hasFiltersDivision: boolean;
    clearFilters: boolean;
}
export interface IDetailsListCustomColumnsResultsProp {
    parentCallback: any;
    items?: IResult[];
    searchTerms?: any;
    order?: string;
    showOrganization: boolean;
    showDepartment: boolean;
    showDivision: boolean;
}
export interface IDetailsListCustomColumnsResultsState {
    sortedItems: any;
    columns: IColumn[];
    searchTerms?: any;
    order?: string;
    showOrganization: boolean;
    showDepartment: boolean;
    showDivision: boolean;
}
export interface IFilterPanelProps {
    parentCallback?: any;
    showPanel: boolean;
    filters: string;
    clearFilters: boolean;
}
export interface IFilterPanelState {
    showPanel: boolean;
    hasChoiceData: boolean;
    filters: string;
    filtersOrganization: any;
    filtersDepartment: any;
    filtersDivision: any;
    clearFilters: boolean;
}
export interface IDropdownControlledMultiState {
    selectedItems: string[];
}
export interface IDropdownControlledMultiProps {
    choices?: any;
    label: string;
    placeholder: string;
    onChange: any;
}
export declare class DropdownControlledMulti extends React.Component<IDropdownControlledMultiProps, IDropdownControlledMultiState> {
    constructor(props: any);
    render(): JSX.Element;
    private _onChange;
}
export declare class ContactCard extends React.Component<IContactCardProps, IContactCardState> {
    constructor(props: any);
    componentDidUpdate(previousProps: IContactCardProps, previousState: IContactCardState): void;
    render(): JSX.Element;
}
export declare class ContactCardGrid extends React.Component<IContactCardGridProps, IContactCardGridState> {
    constructor(props: any);
    render(): JSX.Element;
    componentDidMount(): void;
    componentDidUpdate(previousProps: IContactCardGridProps, previousState: IContactCardGridState): void;
    private _getKey;
}
export declare class DetailsListCustomColumnsResults extends React.Component<IDetailsListCustomColumnsResultsProp, IDetailsListCustomColumnsResultsState> {
    constructor(props: any);
    render(): JSX.Element;
    sendData: (order: any) => void;
    componentDidMount(): void;
    componentDidUpdate(previousProps: IDetailsListCustomColumnsResultsProp, previousState: IDetailsListCustomColumnsResultsState): void;
    _renderItemColumn(item: IResult, index: number, column: IColumn, searchTerms: any): JSX.Element | "";
    private _onColumnClick;
    private _onColumnHeaderContextMenu;
    private _onItemInvoked;
}
export declare class FilterPanel extends React.Component<IFilterPanelProps, IFilterPanelState> {
    constructor(props: any);
    componentDidMount(): void;
    componentDidUpdate(previousProps: IFilterPanelProps, previousState: IFilterPanelState): void;
    availOrganizations: any[];
    availOrganizationsObject: any[];
    availDepartments: any[];
    availDepartmentsObject: any[];
    availDivisions: any[];
    availDivisionsObject: any[];
    sortDropdowns(a: any, b: any): 1 | -1;
    getRESTResults(): void;
    sendData: (showPanel: any, filters: any, hasFiltersOrganization: any, hasFiltersDepartment: any, hasFiltersDivision: any, clearFilters: any) => void;
    private _showPanel;
    private _hidePanel;
    private _applyFilters;
    private _clearFilters;
    private _onRenderFooterContent;
    private _onFilterChangeOrganization;
    private _onFilterChangeDepartment;
    private _onFilterChangeDivision;
    render(): JSX.Element;
}
export declare class CommandBarSearchControls extends React.Component<ICommandBarSearchControlsProps, ICommandBarSearchControlsState> {
    constructor(props: any);
    componentDidMount(): void;
    componentDidUpdate(previousProps: ICommandBarSearchControlsProps, previousState: ICommandBarSearchControlsState): void;
    sendData: (boolVal: any, view: any, order: any, size: any, showPanel: any, filters: any, hasFiltersOrganization: any, hasFiltersDepartment: any, hasFiltersDivision: any, clearFilters: any) => void;
    handleFilterClick: () => void;
    handleSortTilesClick: (orderClicked: any) => void;
    handleViewTilesClick: () => void;
    handleViewListClick: () => void;
    handleTileSizeClick: (sizeClicked: any) => void;
    callbackFromFilterPanelToCommandBar: (showPanel: any, filters: any, hasFiltersOrganization: any, hasFiltersDepartment: any, hasFiltersDivision: any, clearFilters: any) => void;
    private getItems;
    private getFarItems;
    render(): JSX.Element;
}
export declare class ContactSearchBox extends React.Component<IContactSearchBoxProps, IContactSearchBoxState> {
    constructor(props: any);
    componentDidMount(): void;
    componentDidUpdate(previousProps: IContactSearchBoxProps, previousState: IContactSearchBoxState): void;
    sendData: (boolVal: any, childData: any, searchTerms: any, view: any, order: any, size: any, showPanel: any, filters: any, hasFiltersOrganization: any, hasFiltersDepartment: any, hasFiltersDivision: any, clearFilters: any) => void;
    handleChange: ((e: any) => void) & _.Cancelable;
    getRESTResults(e: any): void;
    handleClear(e: any): void;
    callbackFromCommandBarToSearchBox: (boolVal: any, view: any, order: any, size: any, showPanel: any, filters: any, hasFiltersOrganization: any, hasFiltersDepartment: any, hasFiltersDivision: any, clearFilters: any) => void;
    render(): JSX.Element;
}
export declare class MainApp extends React.Component<IMainAppProps, IMainAppState> {
    constructor(props: any);
    componentDidMount(): void;
    componentDidUpdate(previousProps: IMainAppProps, previousState: IMainAppState): void;
    componentWillUnmount(): void;
    callbackFromSearchBoxToMainApp: (boolVal: any, childData: any, searchTerms: any, view: any, order: any, size: any, showPanel: any, filters: any, hasFiltersOrganization: any, hasFiltersDepartment: any, hasFiltersDivision: any, clearFilters: any) => void;
    callbackFromDetailsListToMainApp: (order: any) => void;
    render(): JSX.Element;
}
export default class PhoneListSearchWebPart extends BaseClientSideWebPart<IPhoneListSearchWebPartProps> {
    render(): void;
    protected onDispose(): void;
    protected readonly dataVersion: Version;
    protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration;
}
//# sourceMappingURL=PhoneListSearchWebPart.d.ts.map
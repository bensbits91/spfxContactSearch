var __extends = (this && this.__extends) || (function () {
    var extendStatics = Object.setPrototypeOf ||
        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __assign = (this && this.__assign) || Object.assign || function(t) {
    for (var s, i = 1, n = arguments.length; i < n; i++) {
        s = arguments[i];
        for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
            t[p] = s[p];
    }
    return t;
};
import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { PropertyPaneTextField, PropertyPaneCheckbox, PropertyPaneDropdown } from '@microsoft/sp-property-pane';
import * as strings from 'PhoneListSearchWebPartStrings';
import { SearchBox } from 'office-ui-fabric-react/lib/SearchBox';
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { CommandBar } from 'office-ui-fabric-react/lib/CommandBar';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { buildColumns } from 'office-ui-fabric-react/lib/DetailsList';
import { ShimmeredDetailsList } from 'office-ui-fabric-react/lib/ShimmeredDetailsList';
import { PersonaCoin } from 'office-ui-fabric-react/lib/Persona';
import { SPHttpClient } from '@microsoft/sp-http';
import { Dropdown } from 'office-ui-fabric-react/lib/Dropdown';
import styles from './components/PhoneListSearch.module.scss';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { update } from '@microsoft/sp-lodash-subset';
import { debounce } from 'lodash';
var appContext;
var availOrganizationsObject = [];
var propPaneDepartments = [];
var propPaneDivisions = [];
var ContactCard = /** @class */ (function (_super) {
    __extends(ContactCard, _super);
    function ContactCard(props) {
        var _this = _super.call(this, props) || this;
        _this.state = {
            item: _this.props.item,
            showOrganization: _this.props.showOrganization,
            showDepartment: _this.props.showDepartment,
            showDivision: _this.props.showDivision
        };
        return _this;
    }
    ContactCard.prototype.componentDidUpdate = function (previousProps, previousState) {
        if (previousState.item != this.props.item) {
            this.setState({ item: this.props.item }, function () {
            });
        }
        if (previousState.showOrganization != this.props.showOrganization) {
            this.setState({ showOrganization: this.props.showOrganization }, function () {
            });
        }
        if (previousState.showDepartment != this.props.showDepartment) {
            this.setState({ showDepartment: this.props.showDepartment }, function () {
            });
        }
        if (previousState.showDivision != this.props.showDivision) {
            this.setState({ showDivision: this.props.showDivision }, function () {
            });
        }
    };
    ContactCard.prototype.render = function () {
        var searchTerms = this.props.searchTerms;
        var highlightHits = function (str) {
            for (var _i = 0, searchTerms_1 = searchTerms; _i < searchTerms_1.length; _i++) {
                var term = searchTerms_1[_i];
                var searchTermRegex = new RegExp(term.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'), "ig");
                var searchTermHighlighted = '<span style="background-color:yellow;">$&</span>';
                str = str.replace(searchTermRegex, searchTermHighlighted);
            }
            return str;
        };
        return (React.createElement("div", { key: this.props.item.Id, className: this.props.size == 'large' ? styles.contactItem : [styles.contactItem, styles.small].join(' '), "data-item-id": this.props.item.Id },
            React.createElement("div", { className: styles.contactItemImg },
                React.createElement(Link, { href: "https://delve-gcc.office.com/?p=" + this.props.item.Email + "&v=work", target: "about:blank" },
                    React.createElement(PersonaCoin, { text: this.props.item.FirstName != null ? this.props.item.FirstName + ' ' + this.props.item.Title : this.props.item.Title, coinSize: this.props.size == 'large' ? 100 : 50, showInitialsUntilImageLoads: true }))),
            React.createElement("div", { className: styles.contactItemDetails },
                React.createElement("div", { className: styles.padBottom },
                    React.createElement(Link, { href: "https://delve-gcc.office.com/?p=" + this.props.item.Email + "&v=work", target: "about:blank" },
                        React.createElement("div", { className: [styles.contactItemFullName, styles.contactItemFieldBody].join(' '), dangerouslySetInnerHTML: {
                                __html: highlightHits(this.props.item.FirstName != null ? this.props.item.FirstName + ' ' + this.props.item.Title : this.props.item.Title)
                            } })),
                    this.props.item.JobTitle != null
                        ? React.createElement("div", { className: styles.contactItemFieldBody, dangerouslySetInnerHTML: { __html: highlightHits(this.props.item.JobTitle) } })
                        : ''),
                React.createElement("div", { className: styles.padBottom },
                    this.props.item.Organization != null && this.state.showOrganization
                        ? React.createElement("div", { className: styles.contactItemFieldBody, dangerouslySetInnerHTML: { __html: highlightHits(this.props.item.Organization) } })
                        : '',
                    this.props.item.Company != null && this.state.showDepartment
                        ? React.createElement("div", null,
                            React.createElement("span", { className: styles.contactItemFieldLabel }, "Department: "),
                            React.createElement("span", { className: styles.contactItemFieldBody, dangerouslySetInnerHTML: { __html: highlightHits(this.props.item.Company) } }))
                        : '',
                    this.props.item.Division != null && this.state.showDivision
                        ? React.createElement("div", null,
                            React.createElement("span", { className: styles.contactItemFieldLabel }, "Division: "),
                            React.createElement("span", { className: styles.contactItemFieldBody, dangerouslySetInnerHTML: { __html: highlightHits(this.props.item.Division) } }))
                        : '',
                    this.props.item.Program != null
                        ? React.createElement("div", null,
                            React.createElement("span", { className: styles.contactItemFieldLabel }, "Program: "),
                            React.createElement("span", { className: styles.contactItemFieldBody, dangerouslySetInnerHTML: { __html: highlightHits(this.props.item.Program) } }))
                        : ''),
                this.props.item.Email != null
                    ? React.createElement("div", { className: styles.contactItemFieldBody },
                        React.createElement("a", { href: 'mailto:' + this.props.item.Email }, this.props.item.Email))
                    : '',
                this.props.item.WorkPhone != null
                    ? React.createElement("div", { className: styles.contactItemFieldBody }, this.props.item.WorkPhone)
                    : '',
                this.props.item.WorkAddress != null
                    ? React.createElement("div", { className: styles.contactItemFieldBody }, this.props.item.WorkAddress)
                    : '')));
    };
    return ContactCard;
}(React.Component));
export { ContactCard };
var ContactCardGrid = /** @class */ (function (_super) {
    __extends(ContactCardGrid, _super);
    function ContactCardGrid(props) {
        var _this = _super.call(this, props) || this;
        _this.state = {
            items: _this.props.items,
            showOrganization: _this.props.showOrganization,
            showDepartment: _this.props.showDepartment,
            showDivision: _this.props.showDivision,
            size: _this.props.size
        };
        return _this;
    }
    ContactCardGrid.prototype.render = function () {
        var _this = this;
        return (React.createElement("div", null, this.state.items.map(function (item) {
            return (React.createElement(ContactCard, { item: item, searchTerms: _this.props.searchTerms, size: _this.props.size, showOrganization: _this.props.showOrganization, showDepartment: _this.props.showDepartment, showDivision: _this.props.showDivision }));
        })));
    };
    ContactCardGrid.prototype.componentDidUpdate = function (previousProps, previousState) {
        if (previousState.items != this.props.items) {
            this.setState({ items: this.props.items });
        }
    };
    ContactCardGrid.prototype._getKey = function (item, index) {
        return item.key;
    };
    return ContactCardGrid;
}(React.Component));
export { ContactCardGrid };
var DetailsListCustomColumnsResults = /** @class */ (function (_super) {
    __extends(DetailsListCustomColumnsResults, _super);
    function DetailsListCustomColumnsResults(props) {
        var _this = _super.call(this, props) || this;
        _this.sendData = function (order) {
            _this.props.parentCallback(order);
        };
        _this._onColumnClick = function (event, column) {
            var isSortedDescending = column.isSortedDescending;
            if (column.isSorted) {
                isSortedDescending = !isSortedDescending;
            }
            _this.setState({
                order: column.fieldName
            }, function () {
                _this.sendData(_this.state.order);
            });
        };
        _this.state = {
            sortedItems: _this.props.items,
            columns: _buildColumns(_this.props.items, _this.props.showOrganization, _this.props.showDepartment, _this.props.showDivision),
            searchTerms: _this.props.searchTerms,
            order: _this.props.order,
            showOrganization: _this.props.showOrganization,
            showDepartment: _this.props.showDepartment,
            showDivision: _this.props.showDivision
        };
        _this._renderItemColumn = _this._renderItemColumn.bind(_this);
        return _this;
    }
    DetailsListCustomColumnsResults.prototype.render = function () {
        var _a = this.state, sortedItems = _a.sortedItems, columns = _a.columns, searchTerms = _a.searchTerms;
        columns.map(function (column) {
            column.isResizable = true;
            column.name = column.fieldName == 'Company' ? 'Department'
                : column.fieldName == 'Title' ? 'Last Name'
                    : column.fieldName.replace(/([A-Z])/g, ' $1').trim();
        });
        return (React.createElement(ShimmeredDetailsList, { items: sortedItems, setKey: "set", columns: columns, onRenderItemColumn: this._renderItemColumn, onColumnHeaderClick: this._onColumnClick, ariaLabelForSelectionColumn: "Toggle selection", ariaLabelForSelectAllCheckbox: "Toggle selection for all items", checkButtonAriaLabel: "Row checkbox", searchTerms: searchTerms }));
    };
    DetailsListCustomColumnsResults.prototype.componentDidUpdate = function (previousProps, previousState) {
        if (previousState.sortedItems != this.props.items) {
            this.setState({ sortedItems: this.props.items });
        }
        if (previousState.order != this.props.order) {
            this.setState({ order: this.props.order });
        }
        if (previousState.showOrganization != this.props.showOrganization) {
            this.setState({
                showOrganization: this.props.showOrganization,
                columns: _buildColumns(this.props.items, this.props.showOrganization, this.props.showDepartment, this.props.showDivision)
            });
        }
        if (previousState.showDepartment != this.props.showDepartment) {
            this.setState({
                showDepartment: this.props.showDepartment,
                columns: _buildColumns(this.props.items, this.props.showOrganization, this.props.showDepartment, this.props.showDivision)
            });
        }
        if (previousState.showDivision != this.props.showDivision) {
            this.setState({
                showDivision: this.props.showDivision,
                columns: _buildColumns(this.props.items, this.props.showOrganization, this.props.showDepartment, this.props.showDivision)
            });
        }
    };
    DetailsListCustomColumnsResults.prototype._renderItemColumn = function (item, index, column, searchTerms) {
        var searchTermsToHighlight = this.props.searchTerms;
        var highlightHits = function (str) {
            for (var _i = 0, searchTermsToHighlight_1 = searchTermsToHighlight; _i < searchTermsToHighlight_1.length; _i++) {
                var term = searchTermsToHighlight_1[_i];
                var searchTermRegex = new RegExp(term.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'), "ig");
                var searchTermHighlighted = '<span style="background-color:yellow;">$&</span>';
                str = str.replace(searchTermRegex, searchTermHighlighted);
            }
            return str;
        };
        var fieldContent = item[column.fieldName];
        switch (column.key) {
            case 'FirstName':
            case 'Title':
                return (fieldContent != null
                    ? React.createElement(Link, { href: "https://delve-gcc.office.com/?p=" + item.Email + "&v=work", target: "about:blank" },
                        React.createElement("span", { dangerouslySetInnerHTML: { __html: highlightHits(fieldContent) } }))
                    : '');
            case 'JobTitle':
            case 'Organization':
            case 'Company':
            case 'Division':
            case 'Program':
                return (fieldContent != null
                    ? React.createElement("span", { dangerouslySetInnerHTML: { __html: highlightHits(fieldContent) } })
                    : '');
            default:
                return React.createElement("span", null, fieldContent);
        }
    };
    return DetailsListCustomColumnsResults;
}(React.Component));
export { DetailsListCustomColumnsResults };
var FilterPanel = /** @class */ (function (_super) {
    __extends(FilterPanel, _super);
    function FilterPanel(props) {
        var _this = _super.call(this, props) || this;
        _this.sendData = function (showPanel, filters, hasFiltersOrganization, hasFiltersDepartment, hasFiltersDivision, clearFilters) {
            _this.props.parentCallback(showPanel, filters, hasFiltersOrganization, hasFiltersDepartment, hasFiltersDivision, clearFilters);
        };
        _this._showPanel = function () {
            _this.setState({ showPanel: true });
        };
        _this._hidePanel = function () {
            _this.setState({ showPanel: false }, function () { _this.sendData(_this.state.showPanel, _this.state.filters, _this.state.filtersOrganization.length, _this.state.filtersDepartment.length, _this.state.filtersDivision.length, _this.state.clearFilters); });
        };
        _this._applyFilters = function () {
            var restFilters = [];
            var hasFiltersOrganization = false;
            var hasFiltersDepartment = false;
            var hasFiltersDivision = false;
            if (_this.state.prefilter_label_department) {
                if (_this.state.prefilter_label_department != 'NoFilter') {
                    var restFiltersDepartment = "Company eq '" + _this.state.prefilter_label_department.split('&').join('%26') + "'";
                    restFilters.push(restFiltersDepartment);
                    hasFiltersDepartment = true;
                }
            }
            else if (_this.state.filtersDepartment.length) {
                var restFiltersDepartment = "(Company eq '" + _this.state.filtersDepartment.join("' or Company eq '") + "')";
                restFilters.push(restFiltersDepartment);
                hasFiltersDepartment = true;
            }
            if (_this.state.prefilter_label_division) {
                if (_this.state.prefilter_label_division != 'NoFilter') {
                    var restFiltersDivision = "Division eq '" + _this.state.prefilter_label_division.split('&').join('%26') + "'";
                    restFilters.push(restFiltersDivision);
                    hasFiltersDivision = true;
                }
            }
            else if (_this.state.filtersDivision.length) {
                var restFiltersDivision = "(Division eq '" + _this.state.filtersDivision.join("' or Division eq '") + "')";
                restFilters.push(restFiltersDivision);
                hasFiltersDivision = true;
            }
            if (_this.state.filtersOrganization.length) {
                var restFiltersOrganization = "(Organization eq '" + _this.state.filtersOrganization.join("' or Organization eq '") + "')";
                restFilters.push(restFiltersOrganization);
                hasFiltersOrganization = true;
            }
            _this.setState({ filters: restFilters.join(' and ') }, function () {
                _this.sendData(_this.state.showPanel, _this.state.filters, hasFiltersOrganization, hasFiltersDepartment, hasFiltersDivision, _this.state.clearFilters);
            });
        };
        _this._clearFilters = function () {
            _this.setState({
                showPanel: false,
                filters: '',
                filtersOrganization: [],
                filtersDepartment: [],
                filtersDivision: [],
                clearFilters: true
            }, function () {
                _this.sendData(_this.state.showPanel, _this.state.filters, false, false, false, true);
            });
        };
        _this._onRenderFooterContent = function () {
            var applyFilterIcon = { iconName: 'WaitlistConfirmMirrored' };
            var hideFilterIcon = { iconName: 'Hide' };
            var clearFilterIcon = { iconName: 'ClearFilter' };
            return (React.createElement("div", null,
                React.createElement(DefaultButton, { iconProps: applyFilterIcon, onClick: _this._applyFilters }, "Apply"),
                React.createElement(DefaultButton, { iconProps: hideFilterIcon, styles: { root: { marginLeft: 15 } }, onClick: _this._hidePanel }, "Hide"),
                React.createElement(DefaultButton, { iconProps: clearFilterIcon, styles: { root: { marginLeft: 15 } }, onClick: _this._clearFilters }, "Clear")));
        };
        _this._onFilterChangeOrganization = function (e) {
            if (e.target.checked) {
                var newFilters = _this.state.filtersOrganization;
                newFilters.push(e.target.title.split('&').join('%26'));
                _this.setState({
                    filtersOrganization: newFilters
                });
            }
        };
        _this._onFilterChangeDepartment = function (e) {
            if (e.target.checked) {
                var newFilters = _this.state.filtersDepartment;
                newFilters.push(e.target.title.split('&').join('%26'));
                _this.setState({
                    filtersDepartment: newFilters
                });
            }
        };
        _this._onFilterChangeDivision = function (e) {
            if (e.target.checked) {
                var newFilters = _this.state.filtersDivision;
                newFilters.push(e.target.title.split('&').join('%26'));
                _this.setState({
                    filtersDivision: newFilters
                });
            }
        };
        _this.state = {
            showPanel: _this.props.showPanel,
            hasChoiceData: false,
            filters: '',
            filtersOrganization: [],
            filtersDepartment: [],
            filtersDivision: [],
            clearFilters: false,
            prefilter_key_department: '',
            prefilter_key_division: ''
        };
        return _this;
    }
    FilterPanel.prototype.componentDidUpdate = function (previousProps, previousState) {
        var _this = this;
        if (previousState.showPanel != this.props.showPanel) {
            this.setState({ showPanel: this.props.showPanel }, function () {
                _this.sendData(_this.state.showPanel, _this.state.filters, _this.state.filtersOrganization.length, _this.state.filtersDepartment.length, _this.state.filtersDivision.length, _this.state.clearFilters);
            });
        }
        if (previousState.hasChoiceData === false && this.state.hasChoiceData === false) {
            this.setState({ hasChoiceData: true });
        }
        if (previousState.clearFilters != this.props.clearFilters) {
            this.setState({ clearFilters: this.props.clearFilters });
        }
        if (previousState.prefilter_key_department != this.props.prefilter_key_department) {
            this.setState({
                prefilter_key_department: this.props.prefilter_key_department,
                prefilter_label_department: this.props.prefilter_label_department
            }, this._applyFilters);
        }
        if (previousState.prefilter_key_division != this.props.prefilter_key_division) {
            this.setState({
                prefilter_key_division: this.props.prefilter_key_division,
                prefilter_label_division: this.props.prefilter_label_division
            }, this._applyFilters);
        }
    };
    FilterPanel.prototype.render = function () {
        return (React.createElement(Panel, { key: this.state.clearFilters ? 'ReRender' : 'noReRender', isOpen: this.state.showPanel, closeButtonAriaLabel: 'Close', isLightDismiss: true, headerText: 'Filter Contacts', onDismiss: this._hidePanel, onRenderFooterContent: this._onRenderFooterContent, isHiddenOnDismiss: true, isFooterAtBottom: true, type: PanelType.custom, customWidth: '420px' },
            React.createElement(Dropdown, { placeholder: this.state.prefilter_key_department != null
                    && this.state.prefilter_key_department != undefined
                    && this.state.prefilter_key_department != 'NoFilter'
                    ? 'Filtered by ' + this.state.prefilter_label_department
                    : 'Select departments...', label: 'Department', onChange: this._onFilterChangeDepartment, multiSelect: true, options: this.props.departmentOptions, disabled: this.state.prefilter_key_department != null
                    && this.state.prefilter_key_department != undefined
                    && this.state.prefilter_key_department != 'NoFilter', styles: { dropdown: { width: 300 } } }),
            React.createElement(Dropdown, { placeholder: this.state.prefilter_label_division != null && this.state.prefilter_label_division != undefined ? 'Filtered by ' + this.state.prefilter_label_division : 'Select divisions...', label: 'Division', onChange: this._onFilterChangeDivision, multiSelect: true, options: this.props.divisionOptions, styles: { dropdown: { width: 300 } }, disabled: this.state.prefilter_label_division != null && this.state.prefilter_label_division != undefined }),
            React.createElement(Dropdown, { placeholder: 'Select organizations...', label: 'Organization', onChange: this._onFilterChangeOrganization, multiSelect: true, options: availOrganizationsObject, styles: { dropdown: { width: 300 } } })));
    };
    return FilterPanel;
}(React.Component));
export { FilterPanel };
var CommandBarSearchControls = /** @class */ (function (_super) {
    __extends(CommandBarSearchControls, _super);
    function CommandBarSearchControls(props) {
        var _this = _super.call(this, props) || this;
        _this.sendData = function (boolVal, view, order, size, showPanel, filters, hasFiltersOrganization, hasFiltersDepartment, hasFiltersDivision, clearFilters) {
            _this.props.parentCallback(boolVal, view, order, size, showPanel, filters, hasFiltersOrganization, hasFiltersDepartment, hasFiltersDivision, clearFilters);
        };
        _this.handleFilterClick = function () {
            _this.setState({
                showPanel: !_this.state.showPanel
            });
        };
        _this.handleSortTilesClick = function (orderClicked) {
            _this.setState({
                order: orderClicked
            }, function () {
                _this.sendData(true, _this.state.view, _this.state.order, _this.state.size, _this.state.showPanel, _this.state.filters, _this.state.hasFiltersOrganization, _this.state.hasFiltersDepartment, _this.state.hasFiltersDivision, false);
            });
        };
        _this.handleViewTilesClick = function () {
            _this.setState({
                view: 'Tiles'
            }, function () {
                _this.sendData(true, _this.state.view, _this.state.order, _this.state.size, _this.state.showPanel, _this.state.filters, _this.state.hasFiltersOrganization, _this.state.hasFiltersDepartment, _this.state.hasFiltersDivision, false);
            });
        };
        _this.handleViewListClick = function () {
            _this.setState({
                view: 'List'
            }, function () {
                _this.sendData(true, _this.state.view, _this.state.order, _this.state.size, _this.state.showPanel, _this.state.filters, _this.state.hasFiltersOrganization, _this.state.hasFiltersDepartment, _this.state.hasFiltersDivision, false);
            });
        };
        _this.handleTileSizeClick = function (sizeClicked) {
            _this.setState({
                size: sizeClicked
            }, function () {
                _this.sendData(true, _this.state.view, _this.state.order, _this.state.size, _this.state.showPanel, _this.state.filters, _this.state.hasFiltersOrganization, _this.state.hasFiltersDepartment, _this.state.hasFiltersDivision, false);
            });
        };
        _this.callbackFromFilterPanelToCommandBar = function (showPanel, filters, hasFiltersOrganization, hasFiltersDepartment, hasFiltersDivision, clearFilters) {
            _this.setState({
                showPanel: showPanel,
                filters: filters,
                hasFiltersOrganization: hasFiltersOrganization,
                hasFiltersDepartment: hasFiltersDepartment,
                hasFiltersDivision: hasFiltersDivision,
                clearFilters: clearFilters
            }, function () {
                _this.sendData(true, _this.state.view, _this.state.order, _this.state.size, _this.state.showPanel, _this.state.filters, _this.state.hasFiltersOrganization, _this.state.hasFiltersDepartment, _this.state.hasFiltersDivision, clearFilters);
            });
        };
        _this.getItems = function () {
            if (_this.state.view == 'Tiles') {
                return [
                    {
                        key: 'size',
                        name: 'Tile Size',
                        ariaLabel: 'Tile Size',
                        iconProps: {
                            iconName: 'SizeLegacy'
                        },
                        onClick: function () { _this.handleViewListClick(); },
                        subMenuProps: {
                            items: [
                                {
                                    key: 'small',
                                    name: 'Small',
                                    iconProps: {
                                        iconName: 'GridViewSmall'
                                    },
                                    onClick: function () {
                                        _this.handleTileSizeClick('small');
                                    }
                                },
                                {
                                    key: 'large',
                                    name: 'Large',
                                    iconProps: {
                                        iconName: 'GridViewMedium'
                                    },
                                    onClick: function () {
                                        _this.handleTileSizeClick('large');
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
                        onClick: function () { _this.handleViewListClick(); }
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
                    onClick: function () { _this.handleViewTilesClick(); }
                }
            ];
        };
        _this.getFarItems = function () {
            if (_this.state.view == 'Tiles') {
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
                                    onClick: function () {
                                        _this.handleSortTilesClick('FirstName');
                                    }
                                },
                                {
                                    key: 'lastName',
                                    name: 'Last Name',
                                    iconProps: {
                                        iconName: 'UserOptional'
                                    },
                                    onClick: function () {
                                        _this.handleSortTilesClick('Title');
                                    }
                                },
                                {
                                    key: 'organization',
                                    name: 'Organization',
                                    iconProps: {
                                        iconName: 'Org'
                                    },
                                    onClick: function () {
                                        _this.handleSortTilesClick('Organization');
                                    }
                                },
                                {
                                    key: 'department',
                                    name: 'Department',
                                    iconProps: {
                                        iconName: 'Teamwork'
                                    },
                                    onClick: function () {
                                        _this.handleSortTilesClick('Company');
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
                        onClick: function () {
                            _this.handleFilterClick();
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
                    onClick: function () {
                        _this.handleFilterClick();
                    }
                }
            ];
        };
        _this.state = {
            view: _this.props.view,
            order: _this.props.order,
            size: 'small',
            showPanel: _this.props.showPanel,
            filters: _this.props.filters,
            hasFiltersOrganization: false,
            hasFiltersDepartment: false,
            hasFiltersDivision: false,
            clearFilters: _this.props.clearFilters
        };
        _this.handleViewTilesClick = _this.handleViewTilesClick.bind(_this);
        _this.handleViewListClick = _this.handleViewListClick.bind(_this);
        _this.handleSortTilesClick = _this.handleSortTilesClick.bind(_this);
        _this.handleFilterClick = _this.handleFilterClick.bind(_this);
        return _this;
    }
    CommandBarSearchControls.prototype.componentDidUpdate = function (previousProps, previousState) {
        var _this = this;
        if (previousState.filters != this.props.filters) {
            this.setState({ filters: this.props.filters }, function () {
                _this.sendData(true, _this.state.view, _this.state.order, _this.state.size, _this.state.showPanel, _this.state.filters, _this.state.hasFiltersOrganization, _this.state.hasFiltersDepartment, _this.state.hasFiltersDivision, false);
            });
        }
        if (previousState.clearFilters != this.props.clearFilters) {
            this.setState({
                clearFilters: this.props.clearFilters,
                showPanel: false
            }, function () {
                _this.sendData(true, _this.state.view, _this.state.order, _this.state.size, _this.state.showPanel, _this.state.filters, _this.state.hasFiltersOrganization, _this.state.hasFiltersDepartment, _this.state.hasFiltersDivision, _this.state.clearFilters);
            });
        }
    };
    CommandBarSearchControls.prototype.render = function () {
        return (React.createElement("div", null,
            React.createElement(CommandBar, { items: this.getItems(), farItems: this.getFarItems(), ariaLabel: 'Use left and right arrow keys to navigate between commands' }),
            React.createElement(FilterPanel, { parentCallback: this.callbackFromFilterPanelToCommandBar, showPanel: this.state.showPanel, filters: this.state.filters, clearFilters: this.state.clearFilters, prefilter_key_department: this.props.prefilter_key_department, prefilter_key_division: this.props.prefilter_key_division, prefilter_label_department: this.props.prefilter_label_department, prefilter_label_division: this.props.prefilter_label_division, departmentOptions: this.props.departmentOptions, divisionOptions: this.props.divisionOptions })));
    };
    return CommandBarSearchControls;
}(React.Component));
export { CommandBarSearchControls };
var ContactSearchBox = /** @class */ (function (_super) {
    __extends(ContactSearchBox, _super);
    function ContactSearchBox(props) {
        var _this = _super.call(this, props) || this;
        _this.sendData = function (boolVal, childData, searchTerms, view, order, size, showPanel, filters, hasFiltersOrganization, hasFiltersDepartment, hasFiltersDivision, clearFilters) {
            _this.props.parentCallback(boolVal, childData, searchTerms, view, order, size, showPanel, filters, hasFiltersOrganization, hasFiltersDepartment, hasFiltersDivision, clearFilters);
        };
        _this.handleChange = debounce(function (e) {
            if (e.length) {
                _this.getRESTResults(e);
            }
        }, 1000);
        _this.callbackFromCommandBarToSearchBox = function (boolVal, view, order, size, showPanel, filters, hasFiltersOrganization, hasFiltersDepartment, hasFiltersDivision, clearFilters) {
            _this.setState({
                view: view,
                order: order,
                needUpdate: boolVal,
                size: size,
                showPanel: showPanel,
                filters: filters,
                hasFiltersOrganization: hasFiltersOrganization,
                hasFiltersDepartment: hasFiltersDepartment,
                hasFiltersDivision: hasFiltersDivision,
                clearFilters: clearFilters
            }, function () {
                _this.sendData(true, _this.state.items, _this.state.searchTerms, _this.state.view, _this.state.order, _this.state.size, _this.state.showPanel, _this.state.filters, _this.state.hasFiltersOrganization, _this.state.hasFiltersDepartment, _this.state.hasFiltersDivision, clearFilters);
                _this.handleChange(_this.state.searchTerms);
            });
        };
        _this.state = {
            searchTerms: [],
            items: [],
            view: _this.props.view,
            order: _this.props.order,
            needUpdate: false,
            showPanel: false,
            filters: _this.props.filters,
            hasFiltersOrganization: _this.props.hasFiltersOrganization,
            hasFiltersDepartment: _this.props.hasFiltersDepartment,
            hasFiltersDivision: _this.props.hasFiltersDivision,
            showOrganization: _this.props.showOrganization,
            showDepartment: _this.props.showDepartment,
            showDivision: _this.props.showDivision,
            clearFilters: _this.props.clearFilters
        };
        _this.handleChange = _this.handleChange.bind(_this);
        _this.handleClear = _this.handleClear.bind(_this);
        return _this;
    }
    ContactSearchBox.prototype.componentDidUpdate = function (previousProps, previousState) {
        var _this = this;
        if (previousState.order != this.props.order) {
            this.setState({ order: this.props.order, needUpdate: true }, function () {
                if (_this.state.view == 'List') {
                    _this.getRESTResults(_this.state.searchTerms);
                }
            });
        }
        if (previousState.size != this.props.size) {
            this.setState({ size: this.props.size, needUpdate: true }, function () {
            });
        }
        if (previousState.showPanel != this.props.showPanel) {
            this.setState({ showPanel: this.props.showPanel, needUpdate: true }, function () {
            });
        }
        if (previousState.filters != this.state.filters) {
            this.getRESTResults(this.state.searchTerms);
        }
        if (previousState.showOrganization != this.props.showOrganization) {
            this.setState({ showOrganization: this.props.showOrganization, needUpdate: true }, function () {
                _this.getRESTResults(_this.state.searchTerms);
            });
        }
        if (previousState.showDepartment != this.props.showDepartment) {
            this.setState({ showDepartment: this.props.showDepartment, needUpdate: true }, function () {
                _this.getRESTResults(_this.state.searchTerms);
            });
        }
        if (previousState.showDivision != this.props.showDivision) {
            this.setState({ showDivision: this.props.showDivision, needUpdate: true }, function () {
                _this.getRESTResults(_this.state.searchTerms);
            });
        }
        if (previousState.clearFilters != this.props.clearFilters) {
            this.setState({
                clearFilters: this.props.clearFilters
            });
        }
    };
    ContactSearchBox.prototype.getRESTResults = function (e) {
        var _this = this;
        var searchTerms = [];
        var myPromise = new Promise(function (resolve, reject) {
            if (e.constructor === Array) {
                searchTerms = e;
            }
            else {
                searchTerms = e.split(' ');
            }
            var searchFilters = [];
            var searchFields = [
                'Title',
                'FirstName',
                'JobTitle',
                'Program'
            ];
            if (!_this.state.hasFiltersOrganization && _this.state.showOrganization) {
                searchFields.push('Organization');
            }
            if (!_this.state.hasFiltersDepartment && _this.state.showDepartment) {
                searchFields.push('Company');
            }
            if (!_this.state.hasFiltersDivision && _this.state.showDivision) {
                searchFields.push('Division');
            }
            for (var _i = 0, searchTerms_2 = searchTerms; _i < searchTerms_2.length; _i++) {
                var term = searchTerms_2[_i];
                var theseTerms = [];
                for (var _a = 0, searchFields_1 = searchFields; _a < searchFields_1.length; _a++) {
                    var field = searchFields_1[_a];
                    theseTerms.push("substringof('" + term + "'," + field + ")");
                }
                searchFilters.push("(" + theseTerms.join(' or ') + ")");
            }
            var searchSourceUrl = "https://auroragov.sharepoint.com/sites/PhoneList";
            var listName = "EmployeeContactList";
            var select = "$select=Id,Title,FirstName,Email,Company,JobTitle,WorkPhone,WorkAddress,Division,Program,Organization,CellPhone";
            var top = "$top=100";
            var searchBarFilters = "(" + searchFilters.join(' and ') + ")";
            var refiners = _this.state.filters != null && _this.state.filters.length ? _this.state.filters + " and " : '';
            var filter = "$filter=" + refiners + searchBarFilters;
            var sortOrder = '$orderby=' + _this.state.order;
            var requestUrl = searchSourceUrl + "/_api/web/lists/getbytitle('" + listName + "')/items?" + select + "&" + top + "&" + filter + "&" + sortOrder;
            appContext.spHttpClient.get(requestUrl, SPHttpClient.configurations.v1)
                .then(function (response) {
                if (response.ok) {
                    response.json().then(function (responseJSON) {
                        if (responseJSON != null) {
                            var items = responseJSON.value;
                            resolve(items);
                        }
                        reject(new Error('Something went wrong.'));
                    });
                }
            });
        });
        var onResolved = function (items) {
            _this.setState({
                items: items,
                searchTerms: searchTerms,
                view: _this.props.view,
                order: _this.props.order,
                size: _this.props.size,
                hasFiltersOrganization: _this.props.hasFiltersOrganization,
                hasFiltersDepartment: _this.props.hasFiltersDepartment,
                hasFiltersDivision: _this.props.hasFiltersDivision
            }, function () {
                _this.sendData(true, _this.state.items, _this.state.searchTerms, _this.state.view, _this.state.order, _this.state.size, _this.state.showPanel, _this.state.filters, _this.state.hasFiltersOrganization, _this.state.hasFiltersDepartment, _this.state.hasFiltersDivision, false);
            });
        };
        var onRejected = function (error) { return console.log(error); };
        myPromise.then(onResolved, onRejected);
    };
    ContactSearchBox.prototype.handleClear = function (e) {
        var _this = this;
        this.setState({
            items: [],
            searchTerms: '',
            order: ''
        }, function () {
            _this.sendData(true, _this.state.items, _this.state.searchTerms, _this.state.view, _this.state.order, _this.state.size, _this.state.showPanel, _this.state.filters, _this.state.hasFiltersOrganization, _this.state.hasFiltersDepartment, _this.state.hasFiltersDivision, false);
        });
    };
    ContactSearchBox.prototype.render = function () {
        var controls = this.state.items.length
            ? React.createElement(CommandBarSearchControls, { parentCallback: this.callbackFromCommandBarToSearchBox, view: this.state.view, order: this.state.order, showPanel: this.state.showPanel, filters: this.state.filters, clearFilters: this.state.clearFilters, prefilter_key_department: this.props.prefilter_key_department, prefilter_key_division: this.props.prefilter_key_division, prefilter_label_department: this.props.prefilter_label_department, prefilter_label_division: this.props.prefilter_label_division, departmentOptions: this.props.departmentOptions, divisionOptions: this.props.divisionOptions })
            : '';
        return (React.createElement("div", null,
            React.createElement(SearchBox, { underlined: true, placeholder: this.props.searchBoxPlaceholder, onChange: this.handleChange, onClear: this.handleClear }),
            controls));
    };
    return ContactSearchBox;
}(React.Component));
export { ContactSearchBox };
var MainApp = /** @class */ (function (_super) {
    __extends(MainApp, _super);
    function MainApp(props) {
        var _this = _super.call(this, props) || this;
        _this.callbackFromSearchBoxToMainApp = function (boolVal, childData, searchTerms, view, order, size, showPanel, filters, hasFiltersOrganization, hasFiltersDepartment, hasFiltersDivision, clearFilters) {
            _this.setState({
                needUpdate: boolVal,
                items: childData,
                searchTerms: searchTerms,
                view: view,
                order: order,
                size: size,
                showPanel: showPanel,
                filters: filters,
                hasFiltersOrganization: hasFiltersOrganization,
                hasFiltersDepartment: hasFiltersDepartment,
                hasFiltersDivision: hasFiltersDivision,
                clearFilters: clearFilters
            });
        };
        _this.callbackFromDetailsListToMainApp = function (order) {
            _this.setState({
                order: order
            }, function () {
            });
        };
        _this.state = {
            needUpdate: false,
            items: [],
            searchTerms: '',
            view: 'Tiles',
            order: 'FirstName',
            size: 'small',
            showPanel: false,
            filters: '',
            hasFiltersOrganization: false,
            hasFiltersDepartment: false,
            hasFiltersDivision: false,
            clearFilters: false
        };
        _this.callbackFromSearchBoxToMainApp = _this.callbackFromSearchBoxToMainApp.bind(_this);
        return _this;
    }
    MainApp.prototype.render = function () {
        var resultViewElement = this.state.searchTerms.length ?
            this.state.items.length ?
                this.state.view == 'Tiles'
                    ? React.createElement(ContactCardGrid, { items: this.state.items, searchTerms: this.state.searchTerms, size: this.state.size, showOrganization: this.props.showOrganization, showDepartment: this.props.showDepartment, showDivision: this.props.showDivision })
                    : React.createElement(DetailsListCustomColumnsResults, { parentCallback: this.callbackFromDetailsListToMainApp, items: this.state.items, searchTerms: this.state.searchTerms, order: this.state.order, showOrganization: this.props.showOrganization, showDepartment: this.props.showDepartment, showDivision: this.props.showDivision })
                : React.createElement("div", null, this.props.noResultText)
            : React.createElement("div", null, this.props.initialResultText);
        return (React.createElement("div", { id: "appRootWrap" },
            React.createElement("h1", null, this.props.appHeading),
            React.createElement(ContactSearchBox, { parentCallback: this.callbackFromSearchBoxToMainApp, view: this.state.view, order: this.state.order, size: this.state.size, showPanel: this.state.showPanel, filters: this.state.filters, hasFiltersOrganization: this.state.hasFiltersOrganization, hasFiltersDepartment: this.state.hasFiltersDepartment, hasFiltersDivision: this.state.hasFiltersDivision, searchBoxPlaceholder: this.props.searchBoxPlaceholder, showOrganization: this.props.showOrganization, showDepartment: this.props.showDepartment, showDivision: this.props.showDivision, clearFilters: this.state.clearFilters, prefilter_key_department: this.props.prefilter_key_department, prefilter_key_division: this.props.prefilter_key_division, prefilter_label_department: this.props.prefilter_label_department, prefilter_label_division: this.props.prefilter_label_division, departmentOptions: this.props.departmentOptions, divisionOptions: this.props.divisionOptions }),
            resultViewElement));
    };
    return MainApp;
}(React.Component));
export { MainApp };
var PhoneListSearchWebPart = /** @class */ (function (_super) {
    __extends(PhoneListSearchWebPart, _super);
    function PhoneListSearchWebPart() {
        var _this = _super !== null && _super.apply(this, arguments) || this;
        _this.availOrganizations = [];
        return _this;
    }
    PhoneListSearchWebPart.prototype.onInit = function () {
        appContext = this.context;
        this.getOptionsPromise = this.getOptions();
        return this.getOptionsPromise;
    };
    PhoneListSearchWebPart.prototype.sortDropdowns = function (a, b) {
        return (a.key > b.key) ? 1 : -1;
    };
    PhoneListSearchWebPart.prototype.getOptions = function () {
        var _this = this;
        return new Promise(function (resolve2, reject2) {
            var myPromise = new Promise(function (resolve, reject) {
                var searchSourceUrl = "https://auroragov.sharepoint.com/sites/PhoneList";
                var listName = "EmployeeContactList";
                var select = "$select=Company,Division,Organization";
                var top = "$top=500";
                var requestUrl = searchSourceUrl + "/_api/web/lists/getbytitle('" + listName + "')/items?" + select + "&" + top;
                appContext.spHttpClient.get(requestUrl, SPHttpClient.configurations.v1)
                    .then(function (response) {
                    if (response.ok) {
                        response.json().then(function (responseJSON) {
                            if (responseJSON != null) {
                                var items = responseJSON.value;
                                resolve(items);
                            }
                            reject(new Error('Something went wrong.'));
                        });
                    }
                });
            });
            var onResolved = function (items) {
                var departmentsTempArray = [];
                var divisionsTempArray = [];
                update(_this.properties, 'departmentOptions', function () {
                    return [];
                });
                update(_this.properties, 'divisionOptions', function () {
                    return [];
                });
                update(_this.properties, 'organizationOptions', function () {
                    return [];
                });
                items.map(function (item) {
                    if (item.Company != null) {
                        if (departmentsTempArray.indexOf(item.Company) === -1) {
                            departmentsTempArray.push(item.Company);
                            _this.properties.departmentOptions.push({
                                key: item.Company.split(' ').join(''),
                                text: item.Company
                            });
                        }
                    }
                    if (item.Division != null) {
                        if (divisionsTempArray.indexOf(item.Division) === -1) {
                            divisionsTempArray.push(item.Division);
                            _this.properties.divisionOptions.push({
                                key: item.Division.split(' ').join(''),
                                text: item.Division
                            });
                        }
                    }
                    if (item.Organization != null) {
                        if (_this.availOrganizations.indexOf(item.Organization) === -1) {
                            _this.availOrganizations.push(item.Organization);
                            availOrganizationsObject.push({
                                key: item.Organization.split(' ').join(''),
                                text: item.Organization
                            });
                        }
                    }
                });
                _this.properties.departmentOptions.sort(_this.sortDropdowns);
                _this.properties.divisionOptions.sort(_this.sortDropdowns);
                availOrganizationsObject.sort(_this.sortDropdowns);
                var blankOption = {
                    key: 'NoFilter',
                    text: 'No Filter'
                };
                propPaneDepartments = JSON.parse(JSON.stringify(_this.properties.departmentOptions));
                propPaneDepartments.unshift(blankOption);
                propPaneDivisions = JSON.parse(JSON.stringify(_this.properties.divisionOptions));
                propPaneDivisions.unshift(blankOption);
                _this.render();
            };
            var onRejected = function (error) { console.log(error); };
            myPromise.then(onResolved, onRejected);
            resolve2('good to go');
            reject2(new Error('Something went wrong.'));
        });
    };
    PhoneListSearchWebPart.prototype.render = function () {
        var _this = this;
        if (this.properties.departmentOptions) {
            if (this.properties.prefilter_key_department) {
                if (this.properties.prefilter_key_department != 'NoFilter') {
                    var newDeparmentLabel_1 = this.properties.departmentOptions.find(function (obj) { return obj.key == _this.properties.prefilter_key_department; }).text;
                    update(this.properties, 'prefilter_label_department', function () { return newDeparmentLabel_1; });
                }
                else {
                    update(this.properties, 'prefilter_label_department', function () { return ''; });
                }
            }
        }
        if (this.properties.divisionOptions) {
            if (this.properties.prefilter_key_division) {
                if (this.properties.prefilter_key_division != 'NoFilter') {
                    var newDivisionLabel_1 = this.properties.divisionOptions.find(function (obj) { return obj.key == _this.properties.prefilter_key_division; }).text;
                    update(this.properties, 'prefilter_label_division', function () { return newDivisionLabel_1; });
                }
                else {
                    update(this.properties, 'prefilter_label_division', function () { return ''; });
                }
            }
        }
        var element = React.createElement("div", null,
            React.createElement(MainApp, { searchBoxPlaceholder: this.properties.searchBoxPlaceholder, appHeading: this.properties.appHeading, initialResultText: this.properties.initialResultText, noResultText: this.properties.noResultText, showOrganization: this.properties.showOrganization, showDepartment: this.properties.showDepartment, showDivision: this.properties.showDivision, prefilter_key_department: this.properties.prefilter_key_department, prefilter_label_department: this.properties.prefilter_label_department, prefilter_key_division: this.properties.prefilter_key_division, prefilter_label_division: this.properties.prefilter_label_division, departmentOptions: this.properties.departmentOptions, divisionOptions: this.properties.divisionOptions }));
        ReactDom.render(element, this.domElement);
    };
    PhoneListSearchWebPart.prototype.onDispose = function () {
        ReactDom.unmountComponentAtNode(this.domElement);
    };
    Object.defineProperty(PhoneListSearchWebPart.prototype, "dataVersion", {
        get: function () {
            return Version.parse('1.0');
        },
        enumerable: true,
        configurable: true
    });
    PhoneListSearchWebPart.prototype.getPropertyPaneConfiguration = function () {
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
                                PropertyPaneCheckbox('showOrganization', {
                                    text: 'Organization'
                                }),
                                PropertyPaneCheckbox('showDepartment', {
                                    text: 'Department'
                                }),
                                PropertyPaneCheckbox('showDivision', {
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
    };
    return PhoneListSearchWebPart;
}(BaseClientSideWebPart));
export default PhoneListSearchWebPart;
function _buildColumns(items, showOrganization, showDepartment, showDivision) {
    var theColumns = [];
    items.map(function (item) {
        theColumns.push(__assign({ FirstName: item.FirstName, Title: item.Title, JobTitle: item.JobTitle }, showOrganization ? { Organization: item.Organization } : null, showDepartment ? { Company: item.Company } : null, showDivision ? { Division: item.Division } : null, { Program: item.Program, Email: item.Email, WorkPhone: item.WorkPhone, WorkAddress: item.WorkAddress }));
    });
    var columns = buildColumns(theColumns);
    return columns;
}
//# sourceMappingURL=PhoneListSearchWebPart.js.map
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
import { PropertyPaneTextField, PropertyPaneCheckbox } from '@microsoft/sp-property-pane';
import * as strings from 'PhoneListSearchWebPartStrings';
import { mergeStyleSets, getTheme, getFocusStyle } from 'office-ui-fabric-react/lib/Styling';
import { SearchBox } from 'office-ui-fabric-react/lib/SearchBox';
import { DefaultButton } from 'office-ui-fabric-react/lib/Button';
import { CommandBar } from 'office-ui-fabric-react/lib/CommandBar';
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { buildColumns } from 'office-ui-fabric-react/lib/DetailsList';
import { PersonaCoin } from 'office-ui-fabric-react/lib/Persona';
import { SPHttpClient } from '@microsoft/sp-http';
import styles from './components/PhoneListSearch.module.scss';
import { debounce } from 'lodash';
import { Dropdown } from 'office-ui-fabric-react/lib/Dropdown';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { ShimmeredDetailsList } from 'office-ui-fabric-react/lib/ShimmeredDetailsList';
import { graph } from "@pnp/graph";
var theme = getTheme();
var palette = theme.palette, fonts = theme.fonts;
var classNames = mergeStyleSets({
    fileIconHeaderIcon: {
        padding: 0,
        fontSize: '16px'
    },
    fileIconCell: {
        textAlign: 'center',
        selectors: {
            '&:before': {
                content: '.',
                display: 'inline-block',
                verticalAlign: 'middle',
                height: '100%',
                width: '0px',
                visibility: 'hidden'
            }
        }
    },
    fileIconImg: {
        verticalAlign: 'middle',
        maxHeight: '16px',
        maxWidth: '16px'
    },
    controlWrapper: {
        display: 'flex',
        flexWrap: 'wrap'
    },
    exampleToggle: {
        display: 'inline-block',
        marginBottom: '10px',
        marginRight: '30px'
    },
    selectionDetails: {
        marginBottom: '20px'
    },
    listGridExample: {
        overflow: 'hidden',
        fontSize: 0,
        position: 'relative'
    },
    listGridExampleTile: {
        textAlign: 'center',
        outline: 'none',
        position: 'relative',
        float: 'left',
        background: palette.neutralLighter,
        selectors: {
            'focus:after': {
                content: '',
                position: 'absolute',
                left: 2,
                right: 2,
                top: 2,
                bottom: 2,
                boxSizing: 'border-box',
                border: "1px solid " + palette.white
            }
        }
    },
    listGridExampleSizer: {
        paddingBottom: '100%'
    },
    listGridExamplePadder: {
        position: 'absolute',
        left: 2,
        top: 2,
        right: 2,
        bottom: 2
    },
    listGridExampleLabel: {
        background: 'rgba(0, 0, 0, 0.3)',
        color: '#FFFFFF',
        position: 'absolute',
        padding: 10,
        bottom: 0,
        left: 0,
        width: '100%',
        fontSize: fonts.small.fontSize,
        boxSizing: 'border-box'
    },
    listGridExampleImage: {
        position: 'absolute',
        top: 0,
        left: 0,
        width: '100%'
    },
    listGridExampleContent: {
        fontSize: 14,
        left: 0,
        position: 'absolute',
        top: 0,
        width: '100%'
    },
    itemCell: [
        getFocusStyle(theme, { inset: -1 }),
        {
            minHeight: 54,
            padding: 10,
            boxSizing: 'border-box',
            borderBottom: "1px solid #aaa",
            display: 'flex',
            selectors: {
                '&:hover': { background: palette.neutralLight }
            }
        }
    ],
    itemImage: {
        flexShrink: 0
    },
    itemContent: {
        marginLeft: 10,
        overflow: 'hidden',
        flexGrow: 1
    },
    itemName: [
        fonts.xLarge,
        {
            whiteSpace: 'nowrap',
            overflow: 'hidden',
            textOverflow: 'ellipsis'
        }
    ],
    itemIndex: {
        fontSize: fonts.small.fontSize,
        color: palette.neutralTertiary,
        marginBottom: 10
    },
    chevron: {
        alignSelf: 'center',
        marginLeft: 10,
        color: palette.neutralTertiary,
        fontSize: fonts.large.fontSize,
        flexShrink: 0
    }
});
var controlStyles = {
    root: {
        margin: '0 30px 20px 0',
        maxWidth: '300px'
    }
};
var dropdownStyles = {
    dropdown: { width: 300 }
};
var appContext;
var DropdownControlledMulti = /** @class */ (function (_super) {
    __extends(DropdownControlledMulti, _super);
    function DropdownControlledMulti(props) {
        var _this = _super.call(this, props) || this;
        _this._onChange = function (event, item) {
            var newSelectedItems = _this.state.selectedItems.slice();
            if (item.selected) {
                newSelectedItems.push(item.key);
            }
            else {
                var currIndex = newSelectedItems.indexOf(item.key);
                if (currIndex > -1) {
                    newSelectedItems.splice(currIndex, 1);
                }
            }
            _this.setState({
                selectedItems: newSelectedItems
            });
        };
        _this.state = {
            selectedItems: []
        };
        return _this;
    }
    DropdownControlledMulti.prototype.render = function () {
        var selectedItems = this.state.selectedItems;
        var choiceObjects = [];
        this.props.choices.map(function (choice) {
            choiceObjects.push({ key: choice.split(' ').join(''), text: choice });
        });
        return (React.createElement(Dropdown, { placeholder: this.props.placeholder, label: this.props.label, selectedKeys: selectedItems, onChange: this._onChange, multiSelect: true, options: choiceObjects, styles: { dropdown: { width: 300 } } }));
    };
    return DropdownControlledMulti;
}(React.Component));
export { DropdownControlledMulti };
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
    // public componentDidMount() {
    //   console.groupCollapsed('ContactCard -> componentDidMount');
    //   console.log('props', this.props);
    //   console.log('state', this.state);
    //   console.groupEnd();
    // }
    ContactCard.prototype.componentDidUpdate = function (previousProps, previousState) {
        // console.groupCollapsed('ContactCard -> componentDidUpdate');
        // console.groupEnd();
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
            showDivision: _this.props.showDivision
        };
        return _this;
    }
    ContactCardGrid.prototype.render = function () {
        var _this = this;
        return (React.createElement("div", null, this.state.items.map(function (item) {
            return (React.createElement(ContactCard, { item: item, searchTerms: _this.props.searchTerms, size: _this.props.size, showOrganization: _this.props.showOrganization, showDepartment: _this.props.showDepartment, showDivision: _this.props.showDivision }));
        })));
    };
    ContactCardGrid.prototype.componentDidMount = function () {
        console.groupCollapsed('ContactCardGrid -> componentDidMount');
        console.log('props', this.props);
        console.log('state', this.state);
        console.groupEnd();
    };
    ContactCardGrid.prototype.componentDidUpdate = function (previousProps, previousState) {
        console.groupCollapsed('ContactCardGrid -> componentDidUpdate');
        console.log('previousProps', previousProps);
        console.log('props', this.props);
        console.log('previousState', previousState);
        console.log('state', this.state);
        console.groupEnd();
        if (previousState.items != this.props.items) {
            this.setState({ items: this.props.items }, function () {
            });
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
            var columns = _this.state.columns;
            var sortedItems = _this.state.sortedItems;
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
        console.log('DetailsListCustomColumnsResults -> render -> columns', columns);
        return (React.createElement(ShimmeredDetailsList, { items: sortedItems, setKey: "set", columns: columns, onRenderItemColumn: this._renderItemColumn, onColumnHeaderClick: this._onColumnClick, onItemInvoked: this._onItemInvoked, onColumnHeaderContextMenu: this._onColumnHeaderContextMenu, ariaLabelForSelectionColumn: "Toggle selection", ariaLabelForSelectAllCheckbox: "Toggle selection for all items", checkButtonAriaLabel: "Row checkbox", searchTerms: searchTerms }));
    };
    DetailsListCustomColumnsResults.prototype.componentDidMount = function () {
        console.groupCollapsed('DetailsListCustomColumnsResults -> componentDidMount');
        console.log('props', this.props);
        console.log('state', this.state);
        console.groupEnd();
    };
    DetailsListCustomColumnsResults.prototype.componentDidUpdate = function (previousProps, previousState) {
        console.groupCollapsed('DetailsListCustomColumnsResults -> componentDidUpdate');
        console.log('previousProps', previousProps);
        console.log('props', this.props);
        console.log('previousState', previousState);
        console.log('state', this.state);
        console.groupEnd();
        if (previousState.sortedItems != this.props.items) {
            this.setState({ sortedItems: this.props.items }, function () {
            });
        }
        if (previousState.order != this.props.order) {
            this.setState({ order: this.props.order }, function () {
            });
        }
        if (previousState.showOrganization != this.props.showOrganization) {
            this.setState({
                showOrganization: this.props.showOrganization,
                columns: _buildColumns(this.props.items, this.props.showOrganization, this.props.showDepartment, this.props.showDivision)
            }, function () {
            });
        }
        if (previousState.showDepartment != this.props.showDepartment) {
            this.setState({
                showDepartment: this.props.showDepartment,
                columns: _buildColumns(this.props.items, this.props.showOrganization, this.props.showDepartment, this.props.showDivision)
            }, function () {
            });
        }
        if (previousState.showDivision != this.props.showDivision) {
            this.setState({
                showDivision: this.props.showDivision,
                columns: _buildColumns(this.props.items, this.props.showOrganization, this.props.showDepartment, this.props.showDivision)
            }, function () {
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
    DetailsListCustomColumnsResults.prototype._onColumnHeaderContextMenu = function (column, ev) {
        console.log("column " + column.key + " contextmenu opened.");
    };
    DetailsListCustomColumnsResults.prototype._onItemInvoked = function (item, index) {
        alert("Item " + item.name + " at index " + index + " has been invoked.");
    };
    return DetailsListCustomColumnsResults;
}(React.Component));
export { DetailsListCustomColumnsResults };
var FilterPanel = /** @class */ (function (_super) {
    __extends(FilterPanel, _super);
    function FilterPanel(props) {
        var _this = _super.call(this, props) || this;
        _this.availOrganizations = [];
        _this.availOrganizationsObject = [];
        _this.availDepartments = [];
        _this.availDepartmentsObject = [];
        _this.availDivisions = [];
        _this.availDivisionsObject = [];
        _this.sendData = function (showPanel, filters, hasFiltersOrganization, hasFiltersDepartment, hasFiltersDivision, clearFilters) {
            console.groupCollapsed('FilterPanel -> sendData');
            console.log('showPanel, filters, hasFiltersOrganization, hasFiltersDepartment, hasFiltersDivision, clearFilters', showPanel, filters, hasFiltersOrganization, hasFiltersDepartment, hasFiltersDivision, clearFilters);
            console.log('state', _this.state);
            console.groupEnd();
            _this.props.parentCallback(showPanel, filters, hasFiltersOrganization, hasFiltersDepartment, hasFiltersDivision, clearFilters);
        };
        _this._showPanel = function () {
            _this.setState({ showPanel: true });
        };
        _this._hidePanel = function () {
            _this.setState({ showPanel: false }, function () { _this.sendData(_this.state.showPanel, _this.state.filters, _this.state.filtersOrganization.length, _this.state.filtersDepartment.length, _this.state.filtersDivision.length, _this.state.clearFilters); });
        };
        // private _onDismiss = (ev?: React.SyntheticEvent<HTMLElement>) => {
        //   if (!ev) {
        //     console.log('Panel dismissed.');
        //     return;
        //   }
        //   console.log('Close button clicked or light dismissed.');
        //   if (ev.nativeEvent.srcElement && (ev.nativeEvent.srcElement as Element).className.indexOf('ms-Button-icon') !== -1) {
        //     console.log('Close button clicked.');
        //   }
        //   if (ev.nativeEvent.srcElement && (ev.nativeEvent.srcElement as Element).className.indexOf('ms-Overlay') !== -1) {
        //     console.log('Light dismissed.');
        //   }
        //   this._hidePanel();
        //   // this.sendData(false);
        // }
        _this._applyFilters = function () {
            var restFilters = [];
            var hasFiltersOrganization = false;
            var hasFiltersDepartment = false;
            var hasFiltersDivision = false;
            console.groupCollapsed('FilterPanel -> _applyFilters');
            console.log('this.state.filtersOrganization', _this.state.filtersOrganization, _this.state.filtersOrganization.length);
            console.log('this.state.filtersDepartment', _this.state.filtersDepartment, _this.state.filtersDepartment.length);
            console.log('this.state.filtersDivision', _this.state.filtersDivision, _this.state.filtersDivision.length);
            console.groupEnd();
            if (_this.state.filtersOrganization.length) {
                var restFiltersOrganization = "(Organization eq '" + _this.state.filtersOrganization.join("' or Organization eq '") + "')";
                restFilters.push(restFiltersOrganization);
                hasFiltersOrganization = true;
            }
            if (_this.state.filtersDepartment.length) {
                var restFiltersDepartment = "(Company eq '" + _this.state.filtersDepartment.join("' or Company eq '") + "')";
                restFilters.push(restFiltersDepartment);
                hasFiltersDepartment = true;
            }
            if (_this.state.filtersDivision.length) {
                var restFiltersDivision = "(Division eq '" + _this.state.filtersDivision.join("' or Division eq '") + "')";
                restFilters.push(restFiltersDivision);
                hasFiltersDivision = true;
            }
            _this.setState({ filters: restFilters.join(' and ') }, function () {
                _this.sendData(_this.state.showPanel, _this.state.filters, hasFiltersOrganization, hasFiltersDepartment, hasFiltersDivision, _this.state.clearFilters);
            });
        };
        _this._clearFilters = function () {
            console.log('_clearFilters');
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
                newFilters.push(e.target.title);
                _this.setState({
                    filtersOrganization: newFilters
                }, function () {
                });
            }
        };
        _this._onFilterChangeDepartment = function (e) {
            if (e.target.checked) {
                var newFilters = _this.state.filtersDepartment;
                newFilters.push(e.target.title);
                _this.setState({
                    filtersDepartment: newFilters
                }, function () {
                });
            }
        };
        _this._onFilterChangeDivision = function (e) {
            if (e.target.checked) {
                var newFilters = _this.state.filtersDivision;
                newFilters.push(e.target.title);
                _this.setState({
                    filtersDivision: newFilters
                }, function () {
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
            clearFilters: false
        };
        return _this;
    }
    FilterPanel.prototype.componentDidMount = function () {
        console.groupCollapsed('FilterPanel -> componentDidMount');
        console.log('props', this.props);
        console.log('state', this.state);
        console.groupEnd();
    };
    FilterPanel.prototype.componentDidUpdate = function (previousProps, previousState) {
        var _this = this;
        console.groupCollapsed('FilterPanel -> componentDidUpdate');
        console.log('previousProps', previousProps);
        console.log('props', this.props);
        console.log('previousState', previousState);
        console.log('state', this.state);
        console.groupEnd();
        if (previousState.showPanel != this.props.showPanel) {
            this.setState({ showPanel: this.props.showPanel }, function () {
                _this.sendData(_this.state.showPanel, _this.state.filters, _this.state.filtersOrganization.length, _this.state.filtersDepartment.length, _this.state.filtersDivision.length, _this.state.clearFilters);
            });
        }
        if (previousState.hasChoiceData === false && this.state.hasChoiceData === false) {
            this.setState({ hasChoiceData: true }, function () {
                _this.getRESTResults();
            });
        }
        if (previousState.clearFilters != this.props.clearFilters) {
            this.setState({
                clearFilters: this.props.clearFilters /* , needUpdate: true */
            });
        }
    };
    FilterPanel.prototype.sortDropdowns = function (a, b) {
        return /* (a, b) =>  */ (a.text > b.text) ? 1 : -1;
    };
    FilterPanel.prototype.getRESTResults = function () {
        var _this = this;
        var myPromise = new Promise(function (resolve, reject) {
            var searchSourceUrl = "https://auroragov.sharepoint.com/sites/PhoneList";
            var listName = "EmployeeContactList";
            var select = "$select=Company,JobTitle,Division,Program,Organization";
            var top = "$top=5000";
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
            items.map(function (item) {
                if (item.Organization != null) {
                    if (_this.availOrganizations.indexOf(item.Organization) === -1) {
                        _this.availOrganizations.push(item.Organization);
                        _this.availOrganizationsObject.push({
                            key: item.Organization.split(' ').join(''),
                            text: item.Organization
                        });
                    }
                    _this.availOrganizationsObject.sort(_this.sortDropdowns);
                }
                if (item.Company != null) {
                    if (_this.availDepartments.indexOf(item.Company) === -1) {
                        _this.availDepartments.push(item.Company);
                        _this.availDepartmentsObject.push({
                            key: item.Company.split(' ').join(''),
                            text: item.Company
                        });
                    }
                    _this.availDepartmentsObject.sort(_this.sortDropdowns);
                }
                if (item.Division != null) {
                    if (_this.availDivisions.indexOf(item.Division) === -1) {
                        _this.availDivisions.push(item.Division);
                        _this.availDivisionsObject.push({
                            key: item.Division.split(' ').join(''),
                            text: item.Division
                        });
                    }
                    _this.availDivisionsObject.sort(_this.sortDropdowns);
                }
            });
        };
        var onRejected = function (error) { return console.log(error); };
        myPromise.then(onResolved, onRejected);
    };
    FilterPanel.prototype.render = function () {
        return (React.createElement(Panel, { key: this.state.clearFilters ? 'ReRender' : 'noReRender', isOpen: this.state.showPanel, closeButtonAriaLabel: 'Close', isLightDismiss: true, headerText: 'Light Dismiss Panel', onDismiss: this._hidePanel, onRenderFooterContent: this._onRenderFooterContent, isHiddenOnDismiss: true, isFooterAtBottom: true, type: PanelType.custom, customWidth: '420px' },
            React.createElement(Dropdown, { placeholder: 'Select departments...', label: 'Department', onChange: this._onFilterChangeDepartment, multiSelect: true, options: this.availDepartmentsObject, styles: { dropdown: { width: 300 } } }),
            React.createElement(Dropdown, { placeholder: 'Select divisions...', label: 'Division', onChange: this._onFilterChangeDivision, multiSelect: true, options: this.availDivisionsObject, styles: { dropdown: { width: 300 } } }),
            React.createElement(Dropdown, { placeholder: 'Select organizations...', label: 'Organization', onChange: this._onFilterChangeOrganization, multiSelect: true, options: this.availOrganizationsObject, styles: { dropdown: { width: 300 } } })));
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
            console.log('filter clicked');
            _this.setState({
                showPanel: !_this.state.showPanel
            });
        };
        _this.handleSortTilesClick = function (orderClicked) {
            console.log('order clicked');
            _this.setState({
                order: orderClicked
            }, function () {
                _this.sendData(true, _this.state.view, _this.state.order, _this.state.size, _this.state.showPanel, _this.state.filters, _this.state.hasFiltersOrganization, _this.state.hasFiltersDepartment, _this.state.hasFiltersDivision, false);
            });
        };
        _this.handleViewTilesClick = function () {
            console.log('Tiles');
            _this.setState({
                view: 'Tiles'
            }, function () {
                _this.sendData(true, _this.state.view, _this.state.order, _this.state.size, _this.state.showPanel, _this.state.filters, _this.state.hasFiltersOrganization, _this.state.hasFiltersDepartment, _this.state.hasFiltersDivision, false);
            });
        };
        _this.handleViewListClick = function () {
            console.log('List');
            _this.setState({
                view: 'List'
            }, function () {
                _this.sendData(true, _this.state.view, _this.state.order, _this.state.size, _this.state.showPanel, _this.state.filters, _this.state.hasFiltersOrganization, _this.state.hasFiltersDepartment, _this.state.hasFiltersDivision, false);
            });
        };
        _this.handleTileSizeClick = function (sizeClicked) {
            console.log('size clicked');
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
                console.groupCollapsed('CommandBarSearchControls -> callbackFromFilterPanelToCommandBar');
                console.log('showPanel, filters, hasFiltersOrganization, hasFiltersDepartment, hasFiltersDivision, clearFilters', showPanel, filters, hasFiltersOrganization, hasFiltersDepartment, hasFiltersDivision, clearFilters);
                console.log('this.state', _this.state);
                console.groupEnd();
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
        // private getOverlflowItems = () => {
        //   return [
        //     {
        //       key: 'move',
        //       name: 'Move to...',
        //       onClick: () => console.log('Move to'),
        //       iconProps: {
        //         iconName: 'MoveToFolder'
        //       }
        //     },
        //     {
        //       key: 'copy',
        //       name: 'Copy to...',
        //       onClick: () => console.log('Copy to'),
        //       iconProps: {
        //         iconName: 'Copy'
        //       }
        //     },
        //     {
        //       key: 'rename',
        //       name: 'Rename...',
        //       onClick: () => console.log('Rename'),
        //       iconProps: {
        //         iconName: 'Edit'
        //       }
        //     }
        //   ];
        // }
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
            size: 'large',
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
    CommandBarSearchControls.prototype.componentDidMount = function () {
        console.groupCollapsed('CommandBarSearchControls -> componentDidMount');
        console.log('props', this.props);
        console.log('state', this.state);
        console.groupEnd();
    };
    CommandBarSearchControls.prototype.componentDidUpdate = function (previousProps, previousState) {
        var _this = this;
        console.groupCollapsed('CommandBarSearchControls -> componentDidUpdate');
        console.log('previousProps', previousProps);
        console.log('props', this.props);
        console.log('previousState', previousState);
        console.log('state', this.state);
        console.groupEnd();
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
            React.createElement(FilterPanel, { parentCallback: this.callbackFromFilterPanelToCommandBar, showPanel: this.state.showPanel, filters: this.state.filters, clearFilters: this.state.clearFilters })));
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
            else {
                console.log('no data yet');
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
    ContactSearchBox.prototype.componentDidMount = function () {
        console.groupCollapsed('ContactSearchBox -> componentDidMount');
        console.log('props', this.props);
        console.log('state', this.state);
        console.groupEnd();
    };
    ContactSearchBox.prototype.componentDidUpdate = function (previousProps, previousState) {
        var _this = this;
        console.groupCollapsed('ContactSearchBox -> componentDidUpdate');
        console.log('previousProps', previousProps);
        console.log('props', this.props);
        console.log('previousState', previousState);
        console.log('state', this.state);
        console.groupEnd();
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
                clearFilters: this.props.clearFilters /* , needUpdate: true */
            });
        }
    };
    ContactSearchBox.prototype.getRESTResults = function (e) {
        var _this = this;
        console.groupCollapsed('ContactSearchBox -> getRESTResults');
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
            console.log('this.props.hasFiltersOrganization', _this.state.hasFiltersOrganization);
            console.log('this.props.hasFiltersDepartment', _this.state.hasFiltersDepartment);
            console.log('this.props.hasFiltersDivision', _this.state.hasFiltersDivision);
            if (!_this.state.hasFiltersOrganization && _this.state.showOrganization) {
                console.log('no org in refiners, add to searchFields');
                searchFields.push('Organization');
            }
            if (!_this.state.hasFiltersDepartment && _this.state.showDepartment) {
                console.log('no dept in refiners, add to searchFields');
                searchFields.push('Company');
            }
            if (!_this.state.hasFiltersDivision && _this.state.showDivision) {
                console.log('no div in refiners, add to searchFields');
                searchFields.push('Division');
            }
            for (var _i = 0, searchTerms_2 = searchTerms; _i < searchTerms_2.length; _i++) {
                var term = searchTerms_2[_i];
                for (var _a = 0, searchFields_1 = searchFields; _a < searchFields_1.length; _a++) {
                    var field = searchFields_1[_a];
                    searchFilters.push("substringof('" + term + "'," + field + ")");
                }
            }
            var searchSourceUrl = "https://auroragov.sharepoint.com/sites/PhoneList";
            var listName = "EmployeeContactList";
            var select = "$select=Id,Title,FirstName,Email,Company,JobTitle,WorkPhone,WorkAddress,Division,Program,Organization,CellPhone";
            var top = "$top=100";
            var searchBarFilters = "(" + searchFilters.join(' or ') + ")";
            console.log('searchBarFilters', searchBarFilters);
            var refiners = _this.state.filters != null && _this.state.filters.length ? _this.state.filters + " and " : '';
            console.log('refiners', refiners);
            var filter = "$filter=" + refiners + searchBarFilters;
            console.log('filter', filter);
            var sortOrder = '$orderby=' + _this.state.order;
            var requestUrl = searchSourceUrl + "/_api/web/lists/getbytitle('" + listName + "')/items?" + select + "&" + top + "&" + filter + "&" + sortOrder;
            console.log('requestUrl', requestUrl);
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
        console.groupEnd();
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
            ? React.createElement(CommandBarSearchControls, { parentCallback: this.callbackFromCommandBarToSearchBox, view: this.state.view, order: this.state.order, showPanel: this.state.showPanel, filters: this.state.filters, clearFilters: this.state.clearFilters })
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
            }, function () {
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
            size: 'large',
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
    MainApp.prototype.componentDidMount = function () {
        console.group('MainApp -> componentDidMount');
        console.log('this.props', this.props);
        console.log('this.state', this.state);
        console.groupEnd();
        console.log('asdfasdfasdfasdf', graph.users.getById('lhibbs@auroragov.org').photo.toUrl());
        console.log('asdfasdfasdfasdf', graph.users.getById('lhibbs@auroragov.org').photo.toUrlAndQuery());
    };
    MainApp.prototype.componentDidUpdate = function (previousProps, previousState) {
        console.group('MainApp -> componentDidUpdate');
        console.log('previousProps', previousProps);
        console.log('props', this.props);
        console.log('previousState', previousState);
        console.log('state', this.state);
        console.groupEnd();
    };
    MainApp.prototype.componentWillUnmount = function () {
    };
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
            React.createElement(ContactSearchBox, { parentCallback: this.callbackFromSearchBoxToMainApp, view: this.state.view, order: this.state.order, size: this.state.size, showPanel: this.state.showPanel, filters: this.state.filters, hasFiltersOrganization: this.state.hasFiltersOrganization, hasFiltersDepartment: this.state.hasFiltersDepartment, hasFiltersDivision: this.state.hasFiltersDivision, searchBoxPlaceholder: this.props.searchBoxPlaceholder, showOrganization: this.props.showOrganization, showDepartment: this.props.showDepartment, showDivision: this.props.showDivision, clearFilters: this.state.clearFilters }),
            resultViewElement));
    };
    return MainApp;
}(React.Component));
export { MainApp };
var PhoneListSearchWebPart = /** @class */ (function (_super) {
    __extends(PhoneListSearchWebPart, _super);
    function PhoneListSearchWebPart() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    PhoneListSearchWebPart.prototype.render = function () {
        appContext = this.context;
        var element = React.createElement("div", null,
            React.createElement(MainApp, { searchBoxPlaceholder: this.properties.searchBoxPlaceholder, appHeading: this.properties.appHeading, initialResultText: this.properties.initialResultText, noResultText: this.properties.noResultText, showOrganization: this.properties.showOrganization, showDepartment: this.properties.showDepartment, showDivision: this.properties.showDivision }));
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
                            groupName: 'Fields to Show',
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
    console.log('_buildColumns -> items', items);
    var theColumns = [];
    items.map(function (item) {
        theColumns.push(__assign({ FirstName: item.FirstName, Title: item.Title, JobTitle: item.JobTitle }, showOrganization ? { Organization: item.Organization } : null, showDepartment ? { Company: item.Company } : null, showDivision ? { Division: item.Division } : null, { Program: item.Program, Email: item.Email, WorkPhone: item.WorkPhone, WorkAddress: item.WorkAddress }));
    });
    var columns = buildColumns(theColumns);
    return columns;
}
//# sourceMappingURL=PhoneListSearchWebPart.js.map
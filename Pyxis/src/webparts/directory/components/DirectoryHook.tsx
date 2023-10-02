import * as React from 'react';
import { useEffect, useState } from 'react';
import styles from "./Directory.module.scss";
import { PersonaCard } from "./PersonaCard/PersonaCard";
import { RefinersCard } from "./RefinersCard/RefinersCard";
import { spservices } from "../../../SPServices/spservices";
import { IDirectoryState } from "./IDirectoryState";
import * as strings from "DirectoryWebPartStrings";
import {
    Spinner, SpinnerSize, MessageBar, MessageBarType, SearchBox, Icon, Label,
    Pivot, PivotItem, PivotLinkFormat, PivotLinkSize, Dropdown, IDropdownOption, Link
} from "office-ui-fabric-react";
import { Stack, IStackStyles, IStackTokens } from 'office-ui-fabric-react/lib/Stack';
import { debounce } from "throttle-debounce";
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import { ISPServices } from "../../../SPServices/ISPServices";
import { Environment, EnvironmentType } from "@microsoft/sp-core-library";
import { spMockServices } from "../../../SPServices/spMockServices";
import { IDirectoryProps } from './IDirectoryProps';
import Paging from './Pagination/Paging';
import { AadHttpClient, MSGraphClient } from "@microsoft/sp-http";
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { IUserGraphState } from './IUserGraphState';


const slice: any = require('lodash/slice');
const filter: any = require('lodash/filter');
const wrapStackTokens: IStackTokens = { childrenGap: 30 };

let jobTitles = [];
let locations = [];
let departments = [];

const DirectoryHook: React.FC<IDirectoryProps> = (props) => {

    let _services: ISPServices = null;
    let isInitialWithoutQuery: boolean = true;
    if (Environment.type === EnvironmentType.Local) {
        _services = new spMockServices();
    } else {
        _services = new spservices(props.context);
    }
    const [az, setaz] = useState<string[]>([]);
    const [alphaKey, setalphaKey] = useState<string>('All');
    const [departmentKey, setdepartmentKey] = useState<string>('');
    const [jobTitleKey, setjobTitleKey] = useState<string>('');
    const [locationKey, setlocationKey] = useState<string>('');
    const [userGrahArray, setUserGraphArray] = useState<IUserGraphState[]>([]);
    const [state, setstate] = useState<IDirectoryState>({
        users: [],
        isLoading: true,
        errorMessage: "",
        hasError: false,
        indexSelectedKey: "All",
        searchString: "LastName",
        searchText: ""
    });
    const orderOptions: IDropdownOption[] = [
        { key: "FirstName", text: "First Name" },
        { key: "LastName", text: "Last Name" },
        { key: "Department", text: "Department" },
        { key: "Location", text: "Location" },
        { key: "JobTitle", text: "Job Title" }
    ];
    const color = props.context.microsoftTeams ? "white" : "";
    // Paging
    const [pagedItems, setPagedItems] = useState<any[]>([]);
    const [pageSize, setPageSize] = useState<number>(props.pageSize ? props.pageSize : 10);
    const [currentPage, setCurrentPage] = useState<number>(1);

    const _onPageUpdate = async (pageno?: number) => {
        var currentPge = (pageno) ? pageno : currentPage;
        var startItem = ((currentPge - 1) * pageSize);
        var endItem = currentPge * pageSize;
        let filItems = slice(state.users, startItem, endItem);
        setCurrentPage(currentPge);
        setPagedItems(filItems);
    };

    useEffect(() => {
    state.users && state.users.length > 0
         ? state.users.map((user: any) => {

            if (user.JobTitle && jobTitles.indexOf(user.JobTitle) < 0) {
                jobTitles.push(user.JobTitle);
            }


            if (user.OfficeNumber && locations.indexOf(user.OfficeNumber) < 0) {
                locations.push(user.OfficeNumber);
            }
            else if (user.BaseOfficeLocation && locations.indexOf(user.BaseOfficeLocation) < 0) {
                locations.push(user.BaseOfficeLocation);
            }

            if (user.Department && departments.indexOf(user.Department) < 0) {
                departments.push(user.Department);
            }

        }) : [];
    }, [state.users]);

    const diretoryGrid =
        pagedItems && pagedItems.length > 0
            ? pagedItems.map((user: any) => {
                // if (user.JobTitle && jobTitles.indexOf(user.JobTitle) < 0) {
                //     jobTitles.push(user.JobTitle);
                // }

                // if (user.OfficeNumber && locations.indexOf(user.OfficeNumber) < 0) {
                //     locations.push(user.OfficeNumber);
                // }
                // else if (user.BaseOfficeLocation && locations.indexOf(user.BaseOfficeLocation) < 0) {
                //     locations.push(user.BaseOfficeLocation);
                // }

                // if (user.Department && departments.indexOf(user.Department) < 0) {                
                //     departments.push(user.Department);
                // }
                return (
                    <PersonaCard
                        context={props.context}
                        profileProperties={{
                            DisplayName: user.PreferredName,
                            Title: user.JobTitle,
                            PictureUrl: user.PictureURL,
                            Email: user.WorkEmail,
                            Department: user.Department,
                            WorkPhone: user.WorkPhone,
                            Location: user.OfficeNumber
                                ? user.OfficeNumber
                                : user.BaseOfficeLocation
                        }}
                    />
                );
            })
            : [];

    const refinerDepartmentGrid =
        departments && departments.length > 0
            ? departments.sort().map((department: any) => {
                return (
                    <div>
                        <Link onClick={(item) => {
                            departments = [];
                            locations = [];
                            jobTitles = [];
                            let selectedItemText = item.target["innerText"];
                            setstate({ ...state, searchText: "", isLoading: true });
                            setdepartmentKey(selectedItemText);
                            setCurrentPage(1);
                            _searchByAlphabets(false, '', selectedItemText, '');
                        }}>{department}</Link>
                    </div>
                );
            })
            : [];

    const refinerJobTitleGrid =
        jobTitles && jobTitles.length > 0
            ? jobTitles.sort().map((jobTitle: any) => {
                return (
                    <div>
                        <Link onClick={(item) => {
                            departments = [];
                            locations = [];
                            jobTitles = [];
                            let selectedItemText = item.target["innerText"];
                            setstate({ ...state, searchText: "", isLoading: true });
                            setjobTitleKey(selectedItemText);
                            setCurrentPage(1);
                            _searchByAlphabets(false, '', '', selectedItemText);
                        }}>{jobTitle}</Link>
                    </div>
                );
            })
            : [];

    const refinerLocationGrid =
        locations && locations.length > 0
            ? locations.sort().map((location: any) => {
                // Custom Filter
                if (location.toLowerCase().indexOf("connaught place") < 0) {
                    return (
                        <div>
                            <Link onClick={(item) => {
                                departments = [];
                                locations = [];
                                jobTitles = [];
                                let selectedItemText = item.target["innerText"];
                                setstate({ ...state, searchText: "", isLoading: true });
                                setlocationKey(selectedItemText);
                                setCurrentPage(1);
                                _searchByAlphabets(false, selectedItemText, '', '');
                            }}>{location}</Link>
                        </div>
                    );
                }
            })
            : [];

    const _loadAlphabets = () => {
        let alphabets: string[] = [];
        alphabets.push("All");
        for (let i = 65; i < 91; i++) {
            alphabets.push(
                String.fromCharCode(i)
            );
        }
        setaz(alphabets);
    };

    const _alphabetChange = async (item?: PivotItem, ev?: React.MouseEvent<HTMLElement>) => {
        setstate({ ...state, searchText: "", indexSelectedKey: item.props.itemKey, isLoading: true });
        setalphaKey(item.props.itemKey);
        if (item.props.itemKey == 'All') {
            setdepartmentKey('');
            setlocationKey('');
            setjobTitleKey('');
            _searchByAlphabets(true);
        }
        setCurrentPage(1);
    };
    const _searchByAlphabets = async (initialSearch: boolean, locationparam?: string, departmentparam?: string, jobtitleparam?: string) => {
        setstate({ ...state, isLoading: true, searchText: '' });
        let users = null;
        //console.log(userGrahArray);
        if (initialSearch || (`${alphaKey}` == "All" && (!locationKey && !locationparam) && (!jobTitleKey && !jobtitleparam) && (!departmentKey && !departmentparam))) {

            let queryValue = getParameterByName("search");

            if (queryValue && isInitialWithoutQuery) {
                document.querySelectorAll("input[id*=SearchBox]")[0]["value"] = queryValue;
                _searchUsers(queryValue, true);
            }
            else {
                if (props.searchFirstName)
                    users = await _services.searchUsersNew('', `*`, false, userGrahArray);
                else users = await _services.searchUsersNew('', '', true, userGrahArray);
            }

        } else {
            let jobTitle = jobTitleKey ? jobTitleKey : (jobtitleparam ? jobtitleparam : '*');
            let location = locationKey ? locationKey : (locationparam ? locationparam : '*');
            let department = departmentKey ? departmentKey : (departmentparam ? departmentparam : '*');

            let refinerQueryString = `AND (JobTitle:${jobTitle} AND Department:${department} AND (OfficeNumber:${location} OR BaseOfficeLocation:${location}))`;

            if (props.searchFirstName) {
                users = await _services.searchUsersNew('', `FirstName:${alphaKey == "All" ? '' : alphaKey}* ${refinerQueryString}`, false, userGrahArray);
            }
            else if (locationKey || departmentKey || jobTitleKey || locationparam || departmentparam || jobtitleparam) {
                users = await _services.searchUsersNew('', `(FirstName:${alphaKey == "All" ? '' : alphaKey}* OR LastName:${alphaKey == "All" ? '' : alphaKey}*) ${refinerQueryString}`, false, userGrahArray);
            }
            else users = await _services.searchUsersNew(`${alphaKey == "All" ? '' : alphaKey}`, '', true, userGrahArray);
        }
        setstate({
            ...state,
            searchText: '',
            indexSelectedKey: initialSearch ? 'All' : state.indexSelectedKey,
            users:
                users && users.PrimarySearchResults
                    ? users.PrimarySearchResults
                    : null,
            isLoading: false,
            errorMessage: "",
            hasError: false
        });
    };

    const getParameterByName = (name, url = window.location.href) => {
        name = name.replace(/[\[\]]/g, '\\$&');
        var regex = new RegExp('[?&]' + name + '(=([^&#]*)|&|#|$)'),
            results = regex.exec(url);
        if (!results) return null;
        if (!results[2]) return '';
        return decodeURIComponent(results[2].replace(/\+/g, ' '));
    };

    let _searchUsers = async (searchText: string, isExcludeProps: boolean = false) => {
        try {
            setstate({ ...state, searchText: searchText, isLoading: true });
            if (searchText.length > 0) {
                let searchProps: string[] = props.searchProps && props.searchProps.length > 0 ?
                    props.searchProps.split(',') : ['FirstName', 'LastName', 'WorkEmail', 'Department'];
                let qryText: string = '';
                let finalSearchText: string = searchText ? searchText.replace(/ /g, '+') : searchText;
                if (props.clearTextSearchProps || isExcludeProps) {
                    //let tmpCTProps: string[] = props.clearTextSearchProps.indexOf(',') >= 0 ? props.clearTextSearchProps.split(',') : [props.clearTextSearchProps];
                    let tmpCTProps: string[];
                    if (props.clearTextSearchProps) {
                        tmpCTProps = props.clearTextSearchProps.indexOf(',') >= 0 ? props.clearTextSearchProps.split(',') : [props.clearTextSearchProps];
                    }
                    if (isExcludeProps || (tmpCTProps && tmpCTProps.length > 0)) {
                        searchProps.map((srchprop, index) => {
                            let ctPresent: any[] = filter(tmpCTProps, (o) => { return o.toLowerCase() == srchprop.toLowerCase(); });
                            if (ctPresent.length > 0) {
                                if (index == searchProps.length - 1) {
                                    qryText += `${srchprop}:${searchText}*`;
                                } else qryText += `${srchprop}:${searchText}* OR `;
                            } else {
                                if (index == searchProps.length - 1) {
                                    qryText += `${srchprop}:${finalSearchText}*`;
                                } else qryText += `${srchprop}:${finalSearchText}* OR `;
                            }
                        });
                    } else {
                        searchProps.map((srchprop, index) => {
                            if (index == searchProps.length - 1)
                                qryText += `${srchprop}:${finalSearchText}*`;
                            else qryText += `${srchprop}:${finalSearchText}* OR `;
                        });
                    }
                } else {
                    searchProps.map((srchprop, index) => {
                        if (index == searchProps.length - 1)
                            qryText += `${srchprop}:${finalSearchText}*`;
                        else qryText += `${srchprop}:${finalSearchText}* OR `;
                    });
                }
                console.log(qryText);
                const users = await _services.searchUsersNew('', qryText, false, userGrahArray);
                setstate({
                    ...state,
                    searchText: searchText,
                    indexSelectedKey: '0',
                    users:
                        users && users.PrimarySearchResults
                            ? users.PrimarySearchResults
                            : null,
                    isLoading: false,
                    errorMessage: "",
                    hasError: false
                });

                if (!isExcludeProps) {
                    setalphaKey('0');
                }
            } else {
                setstate({ ...state, searchText: '' });
                isInitialWithoutQuery = false;
                _searchByAlphabets(true);
            }
        } catch (err) {
            setstate({ ...state, errorMessage: err.message, hasError: true });
        }
    };

    const _searchBoxChanged = (newvalue: string): void => {
        setCurrentPage(1);
        _searchUsers(newvalue);
    };
    _searchUsers = debounce(1000, _searchUsers);

    const _sortPeople = async (sortField: string) => {
        let _users = [...state.users];
        _users = _users.sort((a: any, b: any) => {
            switch (sortField) {
                // Sorte by FirstName
                case "FirstName":
                    const aFirstName = a.FirstName ? a.FirstName : "";
                    const bFirstName = b.FirstName ? b.FirstName : "";
                    if (aFirstName.toUpperCase() < bFirstName.toUpperCase()) {
                        return -1;
                    }
                    if (aFirstName.toUpperCase() > bFirstName.toUpperCase()) {
                        return 1;
                    }
                    return 0;
                    break;
                // Sort by LastName
                case "LastName":
                    const aLastName = a.LastName ? a.LastName : "";
                    const bLastName = b.LastName ? b.LastName : "";
                    if (aLastName.toUpperCase() < bLastName.toUpperCase()) {
                        return -1;
                    }
                    if (aLastName.toUpperCase() > bLastName.toUpperCase()) {
                        return 1;
                    }
                    return 0;
                    break;
                // Sort by Location
                case "Location":
                    const aBaseOfficeLocation = a.BaseOfficeLocation
                        ? a.BaseOfficeLocation
                        : "";
                    const bBaseOfficeLocation = b.BaseOfficeLocation
                        ? b.BaseOfficeLocation
                        : "";
                    if (
                        aBaseOfficeLocation.toUpperCase() <
                        bBaseOfficeLocation.toUpperCase()
                    ) {
                        return -1;
                    }
                    if (
                        aBaseOfficeLocation.toUpperCase() >
                        bBaseOfficeLocation.toUpperCase()
                    ) {
                        return 1;
                    }
                    return 0;
                    break;
                // Sort by JobTitle
                case "JobTitle":
                    const aJobTitle = a.JobTitle ? a.JobTitle : "";
                    const bJobTitle = b.JobTitle ? b.JobTitle : "";
                    if (aJobTitle.toUpperCase() < bJobTitle.toUpperCase()) {
                        return -1;
                    }
                    if (aJobTitle.toUpperCase() > bJobTitle.toUpperCase()) {
                        return 1;
                    }
                    return 0;
                    break;
                // Sort by Department
                case "Department":
                    const aDepartment = a.Department ? a.Department : "";
                    const bDepartment = b.Department ? b.Department : "";
                    if (aDepartment.toUpperCase() < bDepartment.toUpperCase()) {
                        return -1;
                    }
                    if (aDepartment.toUpperCase() > bDepartment.toUpperCase()) {
                        return 1;
                    }
                    return 0;
                    break;
                default:
                    break;
            }
        });
        setstate({ ...state, users: _users, searchString: sortField });
    };

    useEffect(() => {
        setPageSize(props.pageSize);
        if (state.users) _onPageUpdate();
    }, [state.users, props.pageSize]);

    useEffect(() => {
        if ((alphaKey.length > 0 && alphaKey != "0") || (!departmentKey || !locationKey || !jobTitleKey) && !state.searchText) _searchByAlphabets(false);
    }, [alphaKey]);

    useEffect(() => {
        props.context.aadHttpClientFactory
            .getClient("https://graph.microsoft.com")
            .then((client: AadHttpClient) => {
                // Search for the users with givenName, surname, or displayName equal to the searchFor value
                return client
                    .get(
                        //`https://graph.microsoft.com/v1.0/users?$select=id,displayname,mail,officeLocation,onPremisesDistinguishedName`,
                        `https://graph.microsoft.com/v1.0/users?$select=id,displayname,mail,officeLocation,employeeType,accountEnabled,onPremisesDistinguishedName&$top=999`,
                        AadHttpClient.configurations.v1
                    );
            })
            .then(response => {
                return response.json();
            })
            .then(json => {
                _apUserGraphArray(json);
            })
            .catch(error => {
                console.error(error);
            });

    }, [props]);

    useEffect(() => {
        _loadAlphabets();
        _searchByAlphabets(true);
    }, [userGrahArray])

    const _apUserGraphArray = (json: any): void => {
        let userArray: IUserGraphState[] = [];

        Promise.all(
            json.value.map(x => {

                userArray.push({
                    displayName: x.displayName,
                    id: x.id,
                    mail: x.mail,
                    officeLocation: x.officeLocation,
                    employeeType: x.employeeType,
                    accountEnabled: x.accountEnabled,
                    onPremisesDistinguishedName: x.onPremisesDistinguishedName
                })
            })
        ).then(function () {

            //Filter for Users
            // var removeUserArray = ["Adam Frahm", "Quentin Van Doosselaere", "Steven Black", "Tarquin Wethered", "Toni Martello", "Julia Maurice", "Marina Knight", "Tanya Clark", "Nicole Wangemann", "Ejike Onuchukwu", "Patrik Johnson"];

            // userArray = userArray.filter(
            //     function(e) {
            //       return this.indexOf(e.displayName) < 0;
            //     },
            //     removeUserArray
            // );

            setUserGraphArray(userArray);
        });
    };

    return (
        <div className={styles.directory}>
            <WebPartTitle displayMode={props.displayMode} title={props.title}
                updateProperty={props.updateProperty} />
            <div className={styles.searchBox} style={{ paddingLeft: '20%' }}>
                <SearchBox placeholder={strings.SearchPlaceHolder} className={styles.searchTextBox}
                    onSearch={_searchUsers}
                    value={state.searchText}
                    onChange={_searchBoxChanged} />
                <div>
                    <Pivot className={styles.alphabets} linkFormat={PivotLinkFormat.tabs}
                        selectedKey={state.indexSelectedKey} onLinkClick={_alphabetChange}
                        linkSize={PivotLinkSize.normal} >
                        {az.map((index: string) => {
                            return (
                                <PivotItem headerText={index} itemKey={index} key={index} />
                            );
                        })}
                    </Pivot>
                </div>
            </div>
            {state.isLoading ? (
                <div style={{ marginTop: '10px', paddingLeft: '20%' }}>
                    <Spinner size={SpinnerSize.large} label={strings.LoadingText} />
                </div>
            ) : (
                <>
                    {state.hasError ? (
                        <div style={{ marginTop: '10px', paddingLeft: '20%' }}>
                            <MessageBar messageBarType={MessageBarType.error}>
                                {state.errorMessage}
                            </MessageBar>
                        </div>
                    ) : (
                        <>
                            { alphaKey.toLocaleLowerCase() != "all" && (!pagedItems || pagedItems.length == 0) ? (
                                <div className={styles.noUsers} style={{ paddingLeft: '20%' }}>
                                    <Icon
                                        iconName={"ProfileSearch"}
                                        style={{ fontSize: "54px", color: color }}
                                    />
                                    <Label>
                                        <span style={{ marginLeft: 5, fontSize: "26px", color: color }}>
                                            {strings.DirectoryMessage}
                                        </span>
                                    </Label>
                                </div>
                            ) : (
                                <>
                                    {/* <div style={{ width: '100%', display: 'inline-block' }}>
                                        <Paging
                                            totalItems={state.users.length}
                                            itemsCountPerPage={pageSize}
                                            onPageUpdate={_onPageUpdate}
                                            currentPage={currentPage} />
                                    </div> */}
                                    {/* <div className={styles.dropDownSortBy}>
                                        <Stack horizontal horizontalAlign="center" wrap tokens={wrapStackTokens}>
                                            <Dropdown
                                                placeholder={strings.DropDownPlaceHolderMessage}
                                                label={strings.DropDownPlaceLabelMessage}
                                                options={orderOptions}
                                                selectedKey={state.searchString}
                                                onChange={(ev: any, value: IDropdownOption) => {
                                                    _sortPeople(value.key.toString());
                                                }}
                                                styles={{ dropdown: { width: 200 } }}
                                            />
                                        </Stack>
                                    </div> */}
                                    <div style={{ width: '100%', display: 'inline-block' }}>
                                        <div style={{ width: '22%', display: 'inline-block' }}>
                                            <Stack horizontal horizontalAlign="center" wrap tokens={wrapStackTokens}>
                                                <div style={{ width: '85%', marginLeft: '0px' }}>
                                                    <h4 style={{ marginBottom: '6px' }}>Departments</h4>
                                                    <div className={styles.refinerSection}>
                                                        {refinerDepartmentGrid}
                                                    </div>

                                                    <h4 style={{ marginBottom: '6px', marginTop: '10px' }}>Job Titles</h4>
                                                    <div className={styles.refinerSection}>
                                                        {refinerJobTitleGrid}
                                                    </div>

                                                    <h4 style={{ marginBottom: '6px', marginTop: '10px' }}>Locations</h4>
                                                    <div className={styles.refinerSection}>
                                                        {refinerLocationGrid}
                                                    </div>
                                                </div>
                                            </Stack>
                                        </div>
                                        <div style={{ width: '78%', display: 'inline-block' }}>
                                            <Stack horizontal horizontalAlign="center" wrap tokens={wrapStackTokens}>
                                                <div>
                                                    {diretoryGrid}
                                                </div>
                                            </Stack>
                                        </div>
                                    </div>
                                    <div style={{ width: '100%', display: 'inline-block' }}>
                                        <Paging
                                            totalItems={(state && state.users) ? state.users.length : 0}
                                            itemsCountPerPage={pageSize}
                                            onPageUpdate={_onPageUpdate}
                                            currentPage={currentPage} />
                                    </div>
                                </>
                            )}
                        </>
                    )}
                </>
            )}
        </div>
    );
};

export default DirectoryHook;
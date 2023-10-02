import { WebPartContext } from "@microsoft/sp-webpart-base";
import { graph } from "@pnp/graph";
import { sp, PeoplePickerEntity, ClientPeoplePickerQueryParameters, SearchQuery, SearchResults, SearchProperty, SortDirection } from '@pnp/sp';
import { PrincipalType } from "@pnp/sp/src/sitegroups";
import { isRelativeUrl } from "office-ui-fabric-react";
import { ISPServices } from "./ISPServices";
import { MSGraphClient } from '@microsoft/sp-http';
import { isUndefined } from "lodash";

export class spservices implements ISPServices {
    constructor(private context: WebPartContext) {
        sp.setup({
            spfxContext: this.context
        });
    }

    public async searchUsers(searchString: string, searchFirstName: boolean): Promise<SearchResults> {
        const _search = !searchFirstName ? `LastName:${searchString}*` : `FirstName:${searchString}*`;
        //const _search = !searchFirstName ? `FirstName:${searchString}*` : ``;
        const searchProperties: string[] = ["FirstName", "LastName", "PreferredName", "WorkEmail", "OfficeNumber", "PictureURL", "WorkPhone", "MobilePhone", "JobTitle", "Department", "Skills", "PastProjects", "BaseOfficeLocation", "SPS-UserType", "GroupId", "AccountName"];
        try {
            if (!searchString) return undefined;
            let users = await sp.searchWithCaching(<SearchQuery>{
                Querytext: _search,
                StartRow: 0,
                RowLimit: 500,
                RowsPerPage: 500,
                EnableInterleaving: true,
                SelectProperties: searchProperties,
                SourceId: 'b09a7990-05ea-4af9-81ef-edfab16c4e31',
                //SortList: [{ "Property": "LastName", "Direction": SortDirection.Ascending }],
                SortList: [{ "Property": "FirstName", "Direction": SortDirection.Ascending }],
            });

            let users1 = await sp.search(<SearchQuery>{
                Querytext: _search,
                StartRow: 500,
                RowLimit: 500,
                RowsPerPage: 500,
                EnableInterleaving: true,
                SelectProperties: searchProperties,
                SourceId: 'b09a7990-05ea-4af9-81ef-edfab16c4e31',
                //SortList: [{ "Property": "LastName", "Direction": SortDirection.Ascending }],
                SortList: [{ "Property": "FirstName", "Direction": SortDirection.Ascending }],
            });

            let filteredUsers = users["_primary"].filter(user => (user.AccountName.indexOf("#ext") < 0 && user.FirstName));
            let filteredUsers1 = users1["_primary"].filter(user => (user.AccountName.indexOf("#ext") < 0 && user.FirstName));

            filteredUsers = filteredUsers.concat(filteredUsers1);

            users["_primary"] = filteredUsers;
            return users;

        } catch (error) {
            Promise.reject(error);
        }
    }

    public async _getImageBase64(pictureUrl: string): Promise<string> {
        return new Promise((resolve, reject) => {
            let image = new Image();
            image.addEventListener("load", () => {
                let tempCanvas = document.createElement("canvas");
                tempCanvas.width = image.width,
                    tempCanvas.height = image.height,
                    tempCanvas.getContext("2d").drawImage(image, 0, 0);
                let base64Str;
                try {
                    base64Str = tempCanvas.toDataURL("image/png");
                } catch (e) {
                    return "";
                }
                resolve(base64Str);
            });
            image.src = pictureUrl;
        });
    }

    public async searchUsersNew(searchString: string, srchQry: string, isInitialSearch: boolean, userGrahArray?: any[]): Promise<SearchResults> {
        let qrytext: string = '';
        //console.log(userGrahArray);
        searchString = searchString == "All" ? '*' : searchString;
        if (isInitialSearch) qrytext = `FirstName:${searchString}* OR LastName:${searchString}*`;
        //if (isInitialSearch) qrytext = `FirstName:${searchString}*`;
        else {
            if (srchQry) qrytext = srchQry;
            else {
                if (searchString) qrytext = searchString;
            }
            if (qrytext.length <= 0) qrytext = `*`;
        }
        const searchProperties: string[] = ["FirstName", "LastName", "PreferredName", "WorkEmail", "OfficeNumber", "PictureURL", "WorkPhone", "MobilePhone", "JobTitle", "Department", "Skills", "PastProjects", "BaseOfficeLocation", "SPS-UserType", "GroupId", "AccountName"];
        try {

            let users = await sp.search(<SearchQuery>{
                Querytext: qrytext,
                StartRow: 0,
                RowLimit: 500,
                RowsPerPage: 500,
                EnableInterleaving: true,
                SelectProperties: searchProperties,
                SourceId: 'b09a7990-05ea-4af9-81ef-edfab16c4e31',
                //SortList: [{ "Property": "LastName", "Direction": SortDirection.Ascending }],
                SortList: [{ "Property": "FirstName", "Direction": SortDirection.Ascending }],
            });

            let users1 = await sp.search(<SearchQuery>{
                Querytext: qrytext,
                StartRow: 500,
                RowLimit: 500,
                RowsPerPage: 500,
                EnableInterleaving: true,
                SelectProperties: searchProperties,
                SourceId: 'b09a7990-05ea-4af9-81ef-edfab16c4e31',
                //SortList: [{ "Property": "LastName", "Direction": SortDirection.Ascending }],
                SortList: [{ "Property": "FirstName", "Direction": SortDirection.Ascending }],
            });

            if (users && users.PrimarySearchResults.length > 0) {
                for (let index = 0; index < users.PrimarySearchResults.length; index++) {
                    let user: any = users.PrimarySearchResults[index];
                    if (user.PictureURL) {
                        //user = { ...user, PictureURL: `/_layouts/15/userphoto.aspx?size=M&accountname=${user.WorkEmail}` };
                        user = { ...user, PictureURL: `https://pyxisoncology.sharepoint.com/_layouts/15/userphoto.aspx?size=M&accountname=${user.WorkEmail}` };
                        users.PrimarySearchResults[index] = user;
                    }
                }
            }

            if (users1 && users1.PrimarySearchResults.length > 0) {
                for (let index = 0; index < users1.PrimarySearchResults.length; index++) {
                    let user: any = users1.PrimarySearchResults[index];
                    if (user.PictureURL) {
                        //user = { ...user, PictureURL: `/_layouts/15/userphoto.aspx?size=M&accountname=${user.WorkEmail}` };
                        user = { ...user, PictureURL: `https://pyxisoncology.sharepoint.com/_layouts/15/userphoto.aspx?size=M&accountname=${user.WorkEmail}` };
                        users1.PrimarySearchResults[index] = user;
                    }
                }
            }

            //let filteredUsers = users["_primary"].filter(user => (user.AccountName.indexOf("#ext")<0 && user.FirstName));
            //let filteredUsers = users["_primary"].filter(user => (user.AccountName.indexOf("#ext")<0 && user.FirstName && (user.Department != null) && (user.JobTitle != null)));
            //let filteredUsers = users["_primary"].filter(user => ((user.AccountName.indexOf("#ext") < 0) && user.FirstName && (user.FirstName.indexOf("adm_") < 0)));
            //let filteredUsers1 = users1["_primary"].filter(user => ((user.AccountName.indexOf("#ext") < 0) && user.FirstName && (user.FirstName.indexOf("adm_") < 0)));
            let filteredUsers = users["_primary"].filter(user => ((user.AccountName.indexOf("#ext") < 0) && user.FirstName));
            let filteredUsers1 = users1["_primary"].filter(user => ((user.AccountName.indexOf("#ext") < 0) && user.FirstName));

            filteredUsers = filteredUsers.concat(filteredUsers1);

            //users["_primary"] = filteredUsers;
            //return users;

            if (userGrahArray) {
                if (userGrahArray.length > 0) {
                    // var finaluserGrahArray = userGrahArray.filter(element => {
                    //     return element.onPremisesDistinguishedName !== null;
                    // });

                    var fArray = [];
                    for (var i = 0; i < userGrahArray.length; i++) {
                        //filteredUsers.filter(function(e1){return e1.WorkEmail == "abendin@glenviewcapital.com"})
                        var cUsers = filteredUsers.filter(function (obj) {
                            // Filter - employeeType
                            if ((userGrahArray[i].employeeType == "FTE" || userGrahArray[i].employeeType == "Consultant_CD") && userGrahArray[i].accountEnabled == true) {
                            // Filter - employeeType
                            
                                if (obj.WorkEmail) {
                                    return obj.WorkEmail == userGrahArray[i].mail
                                }
                            }
                        })
                        if (cUsers.length > 0) {
                            //This guy
                            fArray.push(cUsers[0])
                        }
                    }

                    filteredUsers = fArray;
                    if (isUndefined(fArray)) {
                        users["_primary"] = []
                    }
                    else {
                        users["_primary"] = filteredUsers;
                    }

                    console.log(users)
                    return users;
                }
            }

        } catch (error) {
            Promise.reject(error);
        }
    }
}

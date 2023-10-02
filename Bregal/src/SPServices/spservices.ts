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
        //const _search = !searchFirstName ? `LastName:${searchString}*` : `FirstName:${searchString}*`;
        const _search = !searchFirstName ? `FirstName:${searchString}*` : ``;
        const searchProperties: string[] = ["FirstName", "LastName", "PreferredName", "WorkEmail", "OfficeNumber", "PictureURL", "WorkPhone", "MobilePhone", "JobTitle", "Department", "Skills", "PastProjects", "BaseOfficeLocation", "SPS-UserType", "GroupId", "AccountName", "CompanyName"];
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
        //if (isInitialSearch) qrytext = `FirstName:${searchString}* OR LastName:${searchString}*`;

        // Search User By

        //var e = document.getElementById("searchPeopleBy") as HTMLSelectElement;
        //var value = e.options[e.selectedIndex].value;


        // var qrytextVal = "";
        // if (value == "FirstName") {
        //     qrytextVal = `FirstName:${searchString}*`;
        // }
        // else {
        //     qrytextVal = `FirstName:${searchString}*`;
        // }

        // Search User By



        if (isInitialSearch) {
            qrytext = `FirstName:${searchString}*`;
        }
        //if (isInitialSearch) qrytext = qrytextVal;
        else {
            if (srchQry) {
                if(srchQry.indexOf(",") != -1)
                {
                    var qrytext1 = srchQry.split(",")[1];
                    qrytext1 = qrytext1.slice(12);

                    srchQry = srchQry.split(",")[0];

                    if(srchQry.length < 58)
                    {
                        userGrahArray = userGrahArray.filter(e => e.companyName == qrytext1)
                    }
                }
                qrytext = srchQry;

            }
            else {
                if (searchString) {
                    qrytext = searchString;
                }
            }
            if (qrytext.length <= 0) {
                qrytext = `*`;
            }

            // var string1 = qrytext.split(",")[0].length;
            // var string2 = qrytext.split(",")[1].length; 


            // if (string1 <=70 && string2 > 15) {    
            //     var qrytext1 = qrytext.split(",")[1];
            //     qrytext1 = qrytext1.split(':').pop().split(')')[0];

            //     qrytext = qrytext.split(",")[0];
            //     userGrahArray = userGrahArray.filter(e => e.companyName == qrytext1)   
            // }
            // else{
            //     var qrytext1 = qrytext.split(",")[1];
            //     qrytext1 = qrytext1.split(':').pop().split(')')[0];

            //     qrytext = qrytext.split(",")[0];
            //     //userGrahArray = userGrahArray.filter(e => e.companyName == qrytext1)
            // }
        }



        const searchProperties: string[] = ["FirstName", "LastName", "PreferredName", "WorkEmail", "OfficeNumber", "PictureURL", "WorkPhone", "MobilePhone", "JobTitle", "Department", "Skills", "PastProjects", "BaseOfficeLocation", "SPS-UserType", "GroupId", "AccountName", "CompanyName"];
        try {

            let users = await sp.search(<SearchQuery>{
                //Querytext: `FirstName:${searchString}*`,
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
                //Querytext: `FirstName:${searchString}*`,
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
            //if (qrytext.split("CompanyName:")[1] != undefined) {
            //let companyNameAny = qrytext.split("CompanyName:")[1].split(" OR OfficeNumber")[0];
            //let companyNameAny = qrytext.split(",")[1].split(':').pop().split(')')[0];
            // if (qrytext1.length < 15) {
            //     userGrahArray = userGrahArray.filter(e => e.companyName == qrytext1)
            // }
            //}


            if (users && users.PrimarySearchResults.length > 0) {
                for (let index = 0; index < users.PrimarySearchResults.length; index++) {
                    let user: any = users.PrimarySearchResults[index];
                    if (user.PictureURL) {
                        //user = { ...user, PictureURL: `/_layouts/15/userphoto.aspx?size=M&accountname=${user.WorkEmail}` };
                        user = { ...user, PictureURL: `https://bregalnet.sharepoint.com/_layouts/15/userphoto.aspx?size=M&accountname=${user.WorkEmail}` };
                        users.PrimarySearchResults[index] = user;
                    }
                    if (userGrahArray.length > 0) {
                        var combination = userGrahArray.filter(o => o.mail === user.WorkEmail);

                        if (combination.length > 0) {
                            user = { ...user, CompanyName: combination[0].companyName }
                        }
                    }
                    users.PrimarySearchResults[index] = user;
                }
            }

            if (users1 && users1.PrimarySearchResults.length > 0) {
                for (let index = 0; index < users1.PrimarySearchResults.length; index++) {
                    let user: any = users1.PrimarySearchResults[index];
                    if (user.PictureURL) {
                        //user = { ...user, PictureURL: `/_layouts/15/userphoto.aspx?size=M&accountname=${user.WorkEmail}` };
                        user = { ...user, PictureURL: `https://bregalnet.sharepoint.com/_layouts/15/userphoto.aspx?size=M&accountname=${user.WorkEmail}` };
                        users1.PrimarySearchResults[index] = user;
                    }
                    if (userGrahArray.length > 0) {
                        var combination = userGrahArray.filter(o => o.mail === user.WorkEmail);

                        if (combination.length > 0) {
                            user = { ...user, CompanyName: combination[0].companyName }
                        }
                    }
                    users1.PrimarySearchResults[index] = user;
                }
            }

            ///



            //let filteredUsers = users["_primary"].filter(user => (user.AccountName.indexOf("#ext")<0 && user.FirstName));
            //let filteredUsers = users["_primary"].filter(user => (user.AccountName.indexOf("#ext")<0 && user.FirstName && (user.Department != null) && (user.JobTitle != null)));

            // let filteredUsers: any[];
            // let filteredUsers1: any[];

            // if (value == "FirstName") {
            //     filteredUsers = users["_primary"].filter(user => ((user.AccountName.indexOf("#ext") < 0) && user.FirstName && (user.FirstName.indexOf("adm_") < 0)));
            //     filteredUsers1 = users1["_primary"].filter(user => ((user.AccountName.indexOf("#ext") < 0) && user.FirstName && (user.FirstName.indexOf("adm_") < 0)));
            // }
            // else {
            //     //filteredUsers = users["_primary"].filter(user => ((user.AccountName.indexOf("#ext") < 0) && user.Department && (user.FirstName.indexOf("adm_") < 0)));
            //     //filteredUsers1 = users1["_primary"].filter(user => ((user.AccountName.indexOf("#ext") < 0) && user.Department && (user.FirstName.indexOf("adm_") < 0)));
            //     filteredUsers = users["_primary"].filter(user => ((user.AccountName.indexOf("#ext") < 0) && user.CompanyName && (user.FirstName.indexOf("adm_") < 0)));
            //     filteredUsers1 = users1["_primary"].filter(user => ((user.AccountName.indexOf("#ext") < 0) && user.CompanyName && (user.FirstName.indexOf("adm_") < 0)));
            // }

            let filteredUsers = users["_primary"].filter(user => ((user.AccountName.indexOf("#ext") < 0) && user.FirstName && (user.FirstName.indexOf("adm_") < 0)));
            let filteredUsers1 = users1["_primary"].filter(user => ((user.AccountName.indexOf("#ext") < 0) && user.FirstName && (user.FirstName.indexOf("adm_") < 0)));
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
                            if (obj.WorkEmail) {
                                // if (userGrahArray[i].companyName) {
                                //     obj.CompanyName = userGrahArray[i].companyName;
                                // }
                                return obj.WorkEmail == userGrahArray[i].mail;
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

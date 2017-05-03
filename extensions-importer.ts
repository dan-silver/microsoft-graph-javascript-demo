import * as Chance from 'chance';
import { User, WorkbookRange, OpenTypeExtension } from "@microsoft/microsoft-graph-types"

import { GraphClient } from "./GraphHelper";
import { SampleUsers } from "./sample-users";


const driveItemId = "01IN3FJC66ZAPNNUNHFFH2XYV36IFGFCGC";
// https://mod195910-my.sharepoint.com/personal/admin_mod195910_onmicrosoft_com/_layouts/15/WopiFrame.aspx?sourcedoc={d61ec8de-a7d1-4f29-abe2-bbf20a6288c2}&action=editnew
// admin@MOD195910.onmicrosoft.com


async function insertSampleData() {
    const client = await GraphClient();
    const chance = new Chance();


    const sampleData:WorkbookRange = {
        values: SampleUsers.map((user) =>
            [
                user,
                chance.twitter(),
                chance.age(),
                chance.city()
            ]
        )
    }

    return await client
        .api(`/me/drive/items/${driveItemId}/workbook/worksheets/Sheet1/range(address='A1:F9')`)
        .patch(sampleData, (err, res) => {
            debugger;
        });
}

// insertSampleData();
// 







interface UserDetailsExtension extends OpenTypeExtension {
    twitter: string
    age: number
    city: string
}


async function fetchUserDetails() {
    const client = await GraphClient();
    return await client
        .api(`/me/drive/items/${driveItemId}/workbook/worksheets/Sheet1/usedRange`)
        .version(`beta`)
        .get()
        .then((res) => {
            return res.text;
        })
}


async function addUserDetails() {
    const client = await GraphClient();

    const userDetails = await fetchUserDetails();

    let saveExtensionDataPromises = [];

    for (let user of userDetails) {

        let extension:UserDetailsExtension = {
            extensionName: "userDetailsExt",
            twitter: user[1],
            age: user[2],
            city: user[3]
        };


        saveExtensionDataPromises.push(
            client
                .api(`/users/${user[0]}/extensions`)
                .version(`beta`)
                .post(extension)
                .catch((e) => {
                    debugger;
                })
        );
    }

    Promise.all(saveExtensionDataPromises).then(() => {
        console.log("Done saving user data!")
    });
}

addUserDetails();

// View extension data in graph explorer
// https://graph.microsoft.com/beta/users/Adams@MOD195910.onmicrosoft.com?$select=id,displayName,mail,mobilePhone&$expand=extensions





















async function removeAllExtensions() {
    const client = await GraphClient();

    const userDetails = await fetchUserDetails();

    let removeExtensionPromises = [];

    for (let user of SampleUsers) {
        removeExtensionPromises.push(
            client
                .api(`/users/${user}/extensions`)
                .version(`beta`)
                .get()
                .catch((e) => {
                    debugger;
                }).then((res) => {
                    let extensionIds = res['value'].map((extension) => extension.id);
                    let extensionRemovals = [];
                    for (let id of extensionIds) {
                        extensionRemovals.push(
                             client
                                .api(`/users/${user}/extensions/${id}`)
                                .version(`beta`)
                                .delete()
                                .catch((e) => {
                                    debugger;
                                }));
                    }
                    return Promise.all(extensionRemovals);
                })
        );
    
    }

    Promise.all(removeExtensionPromises).then(() => {
        console.log("Done removing extensions")
    });
}

// removeAllExtensions();
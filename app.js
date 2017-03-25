"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
Object.defineProperty(exports, "__esModule", { value: true });
const GraphHelper_1 = require("./GraphHelper");
function queryMicrosoftGraph() {
    return __awaiter(this, void 0, void 0, function* () {
        const client = yield GraphHelper_1.GraphClient();
        return yield client
            .api("/me/trendingAround")
            .version("beta")
            .get()
            .then((res) => {
            debugger;
            console.log(res);
        }).catch((error) => {
            debugger;
        });
    });
}
// queryMicrosoftGraph();
function findUsers() {
    return __awaiter(this, void 0, void 0, function* () {
        const client = yield GraphHelper_1.GraphClient();
        return yield client
            .api("/users")
            .get()
            .then((res) => {
            debugger;
            // console.log(res)
            let users = res.value;
            for (let user of users) {
                console.log(user.displayName, user.mail);
            }
        }).catch((error) => {
            debugger;
        });
    });
}
// findUsers();
let message = {
    subject: "Microsoft Graph TypeScript Sample",
    toRecipients: [{
            emailAddress: {
                address: "dansil@microsoft.com"
            }
        }],
    body: {
        content: "<h1>Microsoft Graph TypeScript Sample</h1>Try modifying the sample",
        contentType: "html" // ***** strongly typed enum - try changing
    }
};
function sendMail() {
    return __awaiter(this, void 0, void 0, function* () {
        const client = yield GraphHelper_1.GraphClient();
        return yield client
            .api("/users/me/sendMail")
            .post({ message })
            .then((res) => {
            console.log("Mail sent!");
        }).catch((error) => {
            debugger;
        });
    });
}
sendMail();
//# sourceMappingURL=app.js.map
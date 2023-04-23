"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
function main(workbook) {
    var _a, _b;
    return __awaiter(this, void 0, void 0, function* () {
        // Replace the {USERNAME} with your Gitub username
        const response = yield fetch('https://api.github.com/users/{USERNAME}/repos');
        const repos = yield response.json();
        const rows = [];
        for (let repo of repos) {
            rows.push([repo.id, repo.name, (_a = repo.license) === null || _a === void 0 ? void 0 : _a.name, (_b = repo.license) === null || _b === void 0 ? void 0 : _b.url]);
        }
        const sheet = workbook.getActiveWorksheet();
        const range = sheet.getRange('A2').getResizedRange(rows.length - 1, rows[0].length - 1);
        range.setValues(rows);
        return;
    });
}

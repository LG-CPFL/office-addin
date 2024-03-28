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
// initialise application
Office.onReady((info) => {
    if (info.host === Office.HostType.Word) {
        const sideloadMsg = document.getElementById("sideload-msg");
        const appBody = document.getElementById("app-body");
        if (sideloadMsg && appBody) {
            sideloadMsg.style.display = "none";
            appBody.style.display = "flex";
            document.addEventListener("DOMContentLoaded", () => {
                attempt(events);
            });
        }
        else {
            console.error("Elements missing");
        }
    }
    else {
        console.error("Host invalid");
    }
});
// error handling
function attempt(fn) {
    return __awaiter(this, void 0, void 0, function* () {
        try {
            yield fn();
        }
        catch (lg) {
            console.error(lg);
        }
    });
}
// create event triggers
function events() {
    // clicking the run button
    const runButton = document.getElementById("run-button");
    if (runButton) {
        runButton.addEventListener("click", () => {
            attempt(main);
        });
    }
}
// execute script
function main() {
    return __awaiter(this, void 0, void 0, function* () {
        return Word.run((context) => __awaiter(this, void 0, void 0, function* () {
            const content = context.document.body;
            content.insertParagraph("Hello There", "End");
            yield context.sync();
        }));
    });
}

// ==UserScript==
// @name         Fill projectflow from timereg
// @namespace    http://netcompany.com/
// @description  Adds a button to ProjectFlow365 that will import registrations from Timereg
// @match        https://ufst.projectflow365.com/*
// @grant        GM_xmlhttpRequest
// @version      0.8
// @connect      timereg.netcompany.com
// @require      https://code.jquery.com/jquery-3.6.0.min.js
// @require      https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.29.4/moment.min.js
// @icon         https://www.google.com/s2/favicons?sz=64&domain=netcompany.com
// ==/UserScript==

let saveButtonSelector = "#cfx-app-PFX_TimeReg_MetaData--7f639013-79d8-4f28-9369-10aed9451fd3-inner > div:nth-child(2) > div > div > div > div > div > div.ms-OverflowSet.ms-CommandBar-primaryCommand.primarySet-209 > div:nth-child(2) > button";
let closeDetailsButtonSelector = "#fluent-default-layer-host > div > div > div > div > div:nth-child(2) > div:nth-child(2) > div > div:nth-child(1) > div > button"
let detailsButtonSelector = "#cfx-app-PFX_Portal_TimeReg--268dadb0-6ea1-4a79-9259-0ec377f1c750-inner > div:nth-child(4) > div > div > div > div > div > div.ms-OverflowSet.ms-CommandBar-secondaryCommand.secondarySet-244 > div:nth-child(1) > button";

const sleep = (delay) => new Promise((resolve) => setTimeout(resolve, delay))

const rolleIdDropdownSelectorN = function (n) {
    let childNum = 6;
    switch (n) {
        case 0:
            break;
        case 1:
            childNum = childNum + 1;
            break;
        default:
            childNum = childNum + 1 + (2 * (n - 1));
    }
    return `#cfx-app-PFX_TimeReg_MetaData--7f639013-79d8-4f28-9369-10aed9451fd3-inner > div.cfx-app-body > div.data-form-module_dataformContainer_JZ7Dk > div > div > div:nth-child(${childNum}) > div > div > div > div > div`;
}

const rolleIdHoursDropdownSelectorN = function (n) {
    let childNum = 8;
    switch (n) {
        case 0:
            throw new Error("rollIdHours 0 does not exist");
        case 1:
            break;
        default:
            childNum = childNum + (2 * (n - 1));
    }
    return `#cfx-app-PFX_TimeReg_MetaData--7f639013-79d8-4f28-9369-10aed9451fd3-inner > div.cfx-app-body > div.data-form-module_dataformContainer_JZ7Dk > div > div > div:nth-child(${childNum}) > div > div > div > div`;
}

function waitForElm(selector) {
    return new Promise(resolve => {
        if (document.querySelector(selector)) {
            return resolve(document.querySelector(selector));
        }

        const observer = new MutationObserver(() => {
            if (document.querySelector(selector)) {
                resolve(document.querySelector(selector));
                observer.disconnect();
            }
        });

        observer.observe(document.body, {
            childList: true,
            subtree: true
        });
    });
}

function waitForSpecificElm(selector, element) {
    return new Promise(resolve => {
        if (element.querySelector(selector)) {
            return resolve(element.querySelector(selector));
        }

        const observer = new MutationObserver(() => {
            if (element.querySelector(selector)) {
                resolve(element.querySelector(selector));
                observer.disconnect();
            }
        });

        observer.observe(element, {
            childList: true,
            subtree: true
        });
    });
}

async function handleRolleIdDropdownAndHours(hours, numRollId) {
    console.log(`Called handleRollId with ${hours} hours and ${numRollId} nummrollid`);
    let rollIdDropdown = document.querySelector(rolleIdDropdownSelectorN(numRollId)).firstChild.firstChild;

    let isAlreadySelected;
    isAlreadySelected = rollIdDropdown.querySelector(".fa.fa-times") != null;

    if (isAlreadySelected) {
        let selectedOptionRemoveButton = rollIdDropdown.querySelector(".fa.fa-times");
        selectedOptionRemoveButton.click();
        await sleep(200);
    }

    if (numRollId > 0) {
        let hoursField = document.querySelector(rolleIdHoursDropdownSelectorN(numRollId)).firstChild.firstChild;
        hoursField.focus();
        hoursField.click();
        await sleep(100);
        hoursField.value = hours.toLocaleString('da-DK');
        const keyboardEvent = new KeyboardEvent('keydown', {
            code: 'Enter',
            key: 'Enter',
            charCode: 13,
            keyCode: 13,
            bubbles: true,
        });
        document.body.dispatchEvent(keyboardEvent);
        await sleep(500);
    }

    rollIdDropdown.lastChild.click();
    let dropdownList = rollIdDropdown.parentNode.children.item(1).firstChild.firstChild.lastChild.getElementsByTagName("li");
    while (dropdownList.length == 0) {
        await sleep(100);
    }
    const element = dropdownList[numRollId]
    element.click();
    await sleep(100);
}

async function testIfWindowHasRollIdDropdown() {
    await waitForElm(saveButtonSelector)
    try {
        let dropdown = document.querySelector(rolleIdDropdownSelectorN(0)).firstChild.firstChild.firstChild.firstChild;
    } catch(e){
        //Is not the expected node
        return false;
    }
    return true;
}

async function closeDetailsWindow() {
    let saveButton = document.querySelector(saveButtonSelector);
    if (saveButton.classList.contains('is-disabled')){
        saveButton = document.querySelector(closeDetailsButtonSelector);
    }
    saveButton.click();
    while (document.querySelector(saveButtonSelector) != null) {
        await new Promise(r => setTimeout(r, 100));
    }
}
async function handleRollId(allRegistrations, isFirstIterationInRow) {

    if (isFirstIterationInRow) {
        let hasRollIds = await testIfWindowHasRollIdDropdown();
        if (!hasRollIds) {
            await closeDetailsWindow();
            return false;
        }
    }
    const sumOfRegistrations = allRegistrations.reduce((a, b) => a + b);
    let currentRollId = 0;
    let currentSumOfRegistrations = 0;

    while (currentSumOfRegistrations < sumOfRegistrations) {
        await handleRolleIdDropdownAndHours(allRegistrations[currentRollId], currentRollId);
        currentSumOfRegistrations += allRegistrations[currentRollId];
        currentRollId += 1;
    }

    await closeDetailsWindow();
    return true;
}

function containsNonNumeric(string) {
    // The regex \D matches any character that's not a digit
    var nonNumericRegex = /\D/;

    // The test() method tests for a match in a string
    return nonNumericRegex.test(string);
}

async function startWait() {
    await waitForElm('.pfx-weeksheet .body-227 .root-228 .primarySet-209');
    $('.pfx-weeksheet .body-227 .root-228 .primarySet-209')
        .append('<div class="ms-OverflowSet-item item-210" role="none"><button id="gmCommDemo" class="ms-Button ms-Button--commandBar ms-CommandBarItem-link root-234" tabindex="0">Fill ProjectFlow from Timereg</button></div>');

    $("#gmCommDemo").click(async function () {
        let year_week = document.querySelector("#cfx-app-PFX_Portal_TimeReg--268dadb0-6ea1-4a79-9259-0ec377f1c750-inner > div.cfx-app-body > div:nth-child(2) > div > div > table > thead > tr.datagrid-module_DataGridGroupHeader_O0qMs > th:nth-child(3)").innerText,
            year = year_week.substr(7, 4), week = year_week.substr(4, 2),
            start_of_week = moment(year + "W" + week).format("YYYY-MM-DD"),
            end_of_week = moment(year + "W" + week).add(6, 'days').format("YYYY-MM-DD"),
            fetchUrl = "https://timereg.netcompany.com/api/Registration/GetRegistrations?startDate=" + start_of_week + "&endDate=" + end_of_week,
            responseJson, projectFlowPsps = [];

        const responseDetails = await new Promise((resolve) => {
            GM_xmlhttpRequest({
                method: 'GET',
                url: fetchUrl,
                onload: function (response) {
                    resolve(response);
                },
            });
        });

        responseJson = JSON.parse(responseDetails.responseText);
        console.log(responseDetails.responseText);
        var insertedSomething = false;

        if (responseJson.Message == "Authorization has been denied for this request.") {
            alert("Missing Authorization. Login to Timereg in another tab");
        }

        var table = document.querySelector("#cfx-app-PFX_Portal_TimeReg--268dadb0-6ea1-4a79-9259-0ec377f1c750-inner > div.cfx-app-body > div:nth-child(2) > div > div > table");
        var rowLength = table.rows.length;
        let deliveryHasMultipleRollIdsMap = new Map();
        // We don't care about the first rows, nor the last summing rows
        for (var i = 3; i < rowLength - 2; i += 1) {
            var row = table.rows[i];
            var projectFlowPsp = row.cells[2].innerText.substr(3, 10);
            if (containsNonNumeric(projectFlowPsp)) continue;
            if (projectFlowPsp.trim() === '') continue;

            for (let delivery of responseJson.RegistrationsGroups) {
                for (let caseRegistration of delivery.CaseRegistrations) {
                    if (caseRegistration.CaseTitle.includes(projectFlowPsp)) {
                        var cellDayStartIndex = 7;
                        let firstIteration = true;
                        for (let registrations of caseRegistration.Registrations) {
                            let allRegistrations = []
                            for (let i = 0; i < registrations.Registrations.length; i++) {
                                allRegistrations[i] = registrations.Registrations[i].Hours;
                            }

                            let hasMoreRollIdsInTimereg = allRegistrations.length > 1;
                            let hourSum = hasMoreRollIdsInTimereg ? allRegistrations.reduce((a, b) => a + b) : allRegistrations[0];

                            if (hourSum > 0) {
                                var cell = row.cells[cellDayStartIndex];

                                if (cell.querySelector(".ms-Icon") != null) continue;
                                cell.focus();
                                cell.click();
                                await waitForSpecificElm('.FitUiControlInput', cell)


                                const enterEvent = new KeyboardEvent('keydown', {
                                    code: 'Enter',
                                    key: 'Enter',
                                    charCode: 13,
                                    keyCode: 13,
                                    bubbles: true,
                                });
                                const tabEvent = new KeyboardEvent('keydown', {
                                    code: 'Tab',
                                    key: 'Tab',
                                    keyCode: 9,
                                    bubbles: true,
                                });
                                let input = cell.lastChild.firstChild.firstChild;
                                input.value = hourSum;
                                input.dispatchEvent(enterEvent);
                                input.dispatchEvent(tabEvent);
                                insertedSomething = true;
                                if (hasMoreRollIdsInTimereg && (firstIteration || deliveryHasMultipleRollIdsMap.get(caseRegistration.CaseTitle))) {
                                    cell.focus();
                                    cell.click();
                                    cell.lastChild.firstChild.firstChild.focus()
                                    await new Promise(r => setTimeout(r, 100));
                                    let detailsButton = document.querySelector(detailsButtonSelector);
                                    let evt = new MouseEvent("click", {
                                        bubbles: true,
                                        cancelable: false,
                                    });
                                    let evt2 = new MouseEvent("mousedown", {
                                        bubbles: true,
                                        cancelable: true,
                                    });
                                    detailsButton.dispatchEvent(evt);
                                    detailsButton.dispatchEvent(evt2);
                                    await waitForElm(saveButtonSelector);
                                    let hasRollId = await handleRollId(allRegistrations, firstIteration);
                                    if (firstIteration) {
                                        deliveryHasMultipleRollIdsMap.set(caseRegistration.CaseTitle, hasRollId);
                                        firstIteration = false;
                                    }
                                }
                            }
                            cellDayStartIndex += 1;
                        }
                    }
                }
            }
        }

        if (!insertedSomething) {
            alert("Nothing was inserted, did you register anything during Week " + week + "?");
        } else {
            setTimeout(function() {
                document.querySelector("#cfx-app-PFX_Portal_TimeReg--268dadb0-6ea1-4a79-9259-0ec377f1c750-inner > div:nth-child(4) > div > div > div > div > div > div.ms-OverflowSet.ms-CommandBar-primaryCommand.primarySet-209 > div:nth-child(2) > button").click()
            }, 1000);
        }
    });
}

startWait();

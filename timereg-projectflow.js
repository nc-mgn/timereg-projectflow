// ==UserScript==
// @name         Fill projectflow from timereg
// @namespace    http://netcompany.com/
// @description  Adds a button to ProjectFlow365 that will import registrations from Timereg
// @match        https://ufst.projectflow365.com/*
// @grant        GM_xmlhttpRequest
// @version      0.3 - Kontraktrolle Id
// @connect      timereg.netcompany.com
// @require      https://code.jquery.com/jquery-3.6.0.min.js
// @require      https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.29.4/moment.min.js
// @icon         https://www.google.com/s2/favicons?sz=64&domain=netcompany.com
// ==/UserScript==

function waitForElm(selector) {
    return new Promise(resolve => {
        if (document.querySelector(selector)) {
            return resolve(document.querySelector(selector));
        }

        const observer = new MutationObserver(mutations => {
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

async function handleRollId() {
    var rollIdDropdown = document.querySelector("#cfx-app-7f639013-79d8-4f28-9369-10aed9451fd3-inner > div.cfx-app-body > div.data-form-module_dataformContainer_JZ7Dk > div > div > div:nth-child(5) > div > div > div > div > div").firstChild.firstChild;

    let isAlreadySelected = false;
    try {
        console.log(rollIdDropdown.firstChild.firstChild)
        console.log(rollIdDropdown.firstChild.firstChild != null);
        isAlreadySelected = rollIdDropdown.firstChild.firstChild != null;
    }catch(e){
        //Fallthrough
    }
    
    if(isAlreadySelected){
        console.log("is true")
    }

    if (!isAlreadySelected) {
        console.log("not true")
    }
    if (isAlreadySelected != true) {
        rollIdDropdown.lastChild.click();
        var dropdownList = rollIdDropdown.parentNode.children.item(1).firstChild.firstChild.lastChild.getElementsByTagName("li");;
        while(dropdownList.length == 0){
            await new Promise(r => setTimeout(r, 100));
        }
        for (var i = 0; i < dropdownList.length; i++) {
            if (dropdownList[i].getAttribute('data-value').includes('A')) {
                dropdownList[i].click();
                break;
            }
        }
    } else {
        await waitForElm("#cfx-app-7f639013-79d8-4f28-9369-10aed9451fd3-inner > div:nth-child(2) > div > div > div > div > div > div.ms-OverflowSet.ms-CommandBar-primaryCommand.primarySet-209 > div:nth-child(2) > button")
    }

    var saveButton = document.querySelector("#cfx-app-7f639013-79d8-4f28-9369-10aed9451fd3-inner > div:nth-child(2) > div > div > div > div > div > div.ms-OverflowSet.ms-CommandBar-primaryCommand.primarySet-209 > div:nth-child(2) > button")
    saveButton.click();
    while (document.querySelector("#cfx-app-7f639013-79d8-4f28-9369-10aed9451fd3-inner > div.cfx-app-body > div.data-form-module_dataformContainer_JZ7Dk > div > div > div:nth-child(5) > div > div > div > div > div") != null){
        await new Promise(r => setTimeout(r, 100));
    }
}

async function startWait() {
    const elm = await waitForElm('.pfx-weeksheet .body-227 .root-228 .primarySet-209');
    $('.pfx-weeksheet .body-227 .root-228 .primarySet-209')
        .append('<div class="ms-OverflowSet-item item-210" role="none"><button id="gmCommDemo" class="ms-Button ms-Button--commandBar ms-CommandBarItem-link root-234" tabindex="0">Fill ProjectFlow from Timereg</button></div>');

    $("#gmCommDemo").click(async function () {
        var year_week = document.querySelector("#cfx-app-268dadb0-6ea1-4a79-9259-0ec377f1c750-inner > div.cfx-app-body > div:nth-child(2) > div > div > table > thead > tr.datagrid-module_DataGridGroupHeader_O0qMs > th:nth-child(3)").innerText;
        var year = year_week.substr(7, 4);
        var week = year_week.substr(4, 2);
        var start_of_week = moment(year + "W" + week).format("YYYY-MM-DD");
        var end_of_week = moment(year + "W" + week).add(6, 'days').format("YYYY-MM-DD");

        var fetchUrl = "https://timereg.netcompany.com/api/Registration/GetRegistrations?startDate=" + start_of_week + "&endDate=" + end_of_week;
        var responseJson;
        var projectFlowPsps = [];

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

        if (responseJson.Message == "Authorization has been denied for this request.") {
            alert("Missing Authorization. Login to Timereg in another tab");
            insertedSomething = true;
        }

        var table = document.querySelector("#cfx-app-268dadb0-6ea1-4a79-9259-0ec377f1c750-inner > div.cfx-app-body > div:nth-child(2) > div > div > table");
        var rowLength = table.rows.length;
        var insertedSomething = false;

        // We don't care about the first rows, nor the last summing rows
        for (var i = 3; i < rowLength - 3; i += 1) {
            var row = table.rows[i];
            var projectFlowPsp = row.cells[2].innerText.substr(3, 10);

            for (let delievery of responseJson.RegistrationsGroups) {
                for (let caseRegistration of delievery.CaseRegistrations) {
                    if (caseRegistration.CaseTitle.includes(projectFlowPsp)) {
                        var cellDayStartIndex = 7;

                        for (let registrations of caseRegistration.Registrations) {
                            for (let registration of registrations.Registrations) {
                                if (registration.Hours > 0) {
                                    var cell = row.cells[cellDayStartIndex];
                                    
                                    cell.focus();
                                    cell.click();

                                    const keyboardEvent = new KeyboardEvent('keydown', {
                                        code: 'Enter',
                                        key: 'Enter',
                                        charCode: 13,
                                        keyCode: 13,
                                        bubbles: true,
                                    });

                                    cell.lastChild.firstChild.firstChild.value = registration.Hours;
                                    cell.lastChild.firstChild.firstChild.dispatchEvent(keyboardEvent);
                                    insertedSomething = true;

                                    cell.focus();
                                    cell.click();
                                    cell.lastChild.firstChild.firstChild.focus()
                                    await new Promise(r => setTimeout(r, 100));
                                    var detailsButton = document.querySelector("#cfx-app-268dadb0-6ea1-4a79-9259-0ec377f1c750-inner > div:nth-child(4) > div > div > div > div > div > div.ms-OverflowSet.ms-CommandBar-secondaryCommand.secondarySet-244 > div:nth-child(1)").firstChild;
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
                                    await waitForElm("#cfx-app-7f639013-79d8-4f28-9369-10aed9451fd3-inner > div.cfx-app-body > div.data-form-module_dataformContainer_JZ7Dk > div > div > div:nth-child(5) > div > div");
                                    await handleRollId();
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
        }
    });
}

startWait();

// ==UserScript==
// @name         Fill projectflow from timereg
// @namespace    http://netcompany.com/
// @description  Adds a button to ProjectFlow365 that will import registrations from Timereg
// @match        https://ufst.projectflow365.com/*
// @grant        GM_xmlhttpRequest
// @version      0.1
// @connect      timereg.netcompany.com
// @require      https://code.jquery.com/jquery-3.6.0.min.js
// @require      https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.29.4/moment.min.js
// @icon         https://www.google.com/s2/favicons?sz=64&domain=netcompany.com
// ==/UserScript==

window.addEventListener('load', function() {
    $("body").prepend ('<button id="gmCommDemo">Fill projectflow from timereg</button>');

    $("#gmCommDemo").click ( function () {
        var year_week = document.querySelector("#cfx-app-268dadb0-6ea1-4a79-9259-0ec377f1c750-inner > div.cfx-app-body > div:nth-child(2) > div > div > table > thead > tr.datagrid-module_DataGridGroupHeader_O0qMs > th:nth-child(3)").innerText;
        var year = year_week.substr(7,4);
        var week = year_week.substr(4,2);
        var start_of_week = moment(year+"W"+week).format("YYYY-MM-DD");
        var end_of_week = moment(year+"W"+week).add(6, 'days').format("YYYY-MM-DD");
        //console.log("Got week " + week);
        //console.log("Start: " + start_of_week + " - End: " + end_of_week);
        var fetchUrl = "https://timereg.netcompany.com/api/Registration/GetRegistrations?startDate="+start_of_week+"&endDate="+end_of_week;
        var responseJson;
        var projectFlowPsps = [];
        GM_xmlhttpRequest ( {
            method:     'GET',
            url:        fetchUrl,
            onload:     function (responseDetails) {
                responseJson = JSON.parse(responseDetails.responseText);
                console.log(responseDetails.responseText);
                if(responseJson.Message == "Authorization has been denied for this request."){
                    alert("Missing Authorization. Login to Timereg in another tab");
                    insertedSomething = true;
                }
                var table = document.querySelector("#cfx-app-268dadb0-6ea1-4a79-9259-0ec377f1c750-inner > div.cfx-app-body > div:nth-child(2) > div > div > table")
                var rowLength = table.rows.length;
                var insertedSomething = false;
                //We don't care about the first rows, nor the last summing rows
                for(var i=3; i<rowLength-3; i+=1){
                    var row = table.rows[i];
                    var projectFlowPsp = row.cells[2].innerText.substr(3,10);
                    jQuery.each(responseJson.RegistrationsGroups, function(i, delievery) {
                        jQuery.each(delievery.CaseRegistrations, function(i, caseRegistration) {
                            if(caseRegistration.CaseTitle.includes(projectFlowPsp)){
                                var cellDayStartIndex = 7;
                                jQuery.each(caseRegistration.Registrations, function(i, registrations) {
                                    var previousCell;
                                    jQuery.each(registrations.Registrations, function(x, registration) {
                                        if(registration.Hours > 0){
                                            //console.log("registration=" + registration);
                                            //console.log("pspCell " + projectFlowPsp);
                                            var cell = row.cells[cellDayStartIndex];
                                            //console.log("cellDayStartIndex=" + cellDayStartIndex);
                                            //console.log("Setting: " + delievery.DeliveryName + " " + caseRegistration.CaseTitle + " " + registrations.Date + " " + registration.Hours);
                                            //cell.click(function(e){ console.log( e ) } );
                                            cell.focus();
                                            cell.click();
                                            const keyboardEvent = new KeyboardEvent('keydown', {
                                                code: 'Enter',
                                                key: 'Enter',
                                                charCode: 13,
                                                keyCode: 13,
                                                //view: window,
                                                bubbles: true
                                            });
                                            cell.lastChild.firstChild.firstChild.value = registration.Hours;
                                            cell.lastChild.firstChild.firstChild.dispatchEvent(keyboardEvent);
                                            insertedSomething = true;
                                        }
                                    });
                                    cellDayStartIndex+=1;
                                });
                            }
                        });
                    });
                }
                if(!insertedSomething){
                    alert("Nothing was inserted, did you register anything during Week " + week + "?");
                }
            }
        } );
    } );
});

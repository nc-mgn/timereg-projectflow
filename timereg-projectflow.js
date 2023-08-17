// ==UserScript==
// @name      Fill projectflow from timereg
// @include    https://ufst.projectflow365.com/*
// @grant     GM_xmlhttpRequest
// @connect   timereg.netcompany.com
// @require   https://code.jquery.com/jquery-3.6.0.min.js
// @require   https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.29.4/moment.min.js
// ==/UserScript==

window.addEventListener('load', function() {
    $("body").prepend ('<button id="gmCommDemo">Fill projectflow from timereg</button>');

    $("#gmCommDemo").click ( function () {
        var week = document.querySelector("#cfx-app-268dadb0-6ea1-4a79-9259-0ec377f1c750-inner > div.TitleArea-module_titleArea_Xc9vN > span").innerText.substr(23,2);
        var start_of_week = moment("2023W"+week).format("YYYY-MM-DD");
        var end_of_week = moment("2023W"+week).add(6, 'days').format("YYYY-MM-DD");
        console.log("Got week " + week);
        console.log("Start: " + start_of_week + " - End: " + end_of_week);
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
                                            console.log("registration=" + registration);
                                            console.log("pspCell " + projectFlowPsp);
                                            var cell = row.cells[cellDayStartIndex];
                                            //console.log("cellDayStartIndex=" + cellDayStartIndex);
                                            console.log("Setting: " + delievery.DeliveryName + " " + caseRegistration.CaseTitle + " " + registrations.Date + " " + registration.Hours);
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
                console.log(projectFlowPsps);
                console.log (
                    "GM_xmlhttpRequest() response is:\n",
                    responseJson.RegistrationsGroups
                );
                if(!insertedSomething){
                    alert("Nothing was inserted, did you register anything during Week " + week + "?");
                }

            }

        } );

        console.log ("Posting message");
    } );
});
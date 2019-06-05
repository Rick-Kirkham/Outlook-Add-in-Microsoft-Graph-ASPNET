/// <reference path="App.js" />
/// <reference path="_officeintellisense.js" />

(function () {
    "use strict";

    var _dlg;
    var redirectTo = "/files/index";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function () {

        $(document).ready(function () {
            app.initialize();

            var authState = {stateKey: stateKey};

            // The stateKey variable must be set on the parent page
            $("#loginO365PopupButton").click(function () {
                var url = "/azureadauth/login?authState=" + encodeURIComponent(JSON.stringify(authState));
                showLoginPopup(url);
            });

        });
    };

    // This handler responds to the success or failure message that the pop-up dialog receives from the identity provider
    // and access token provider.
    function processMessage(arg) {
        console.log("Message received in processMessage: " + JSON.stringify(arg));
        if (arg.message === "success") {
            // We now have a valid access token in the database.
            _dlg.close();

            $("#waitContainer").hide();
            $("#connectContainer").show();
            $("#footer").show();

            window.location.href = redirectTo;
        } else {
            // Something went wrong with authentication or the authorization of the web application.
            _dlg.close();
            app.showNotification("User authentication and application authorization", "Unable to successfully authenticate user or authorize application. Status is " + arg.message);
        }
    }

    // Use the Office dialog API to open a pop-up and display the sign-in page for the identity provider.
    function showLoginPopup(url) {
        $("#connectContainer").hide();
        $("#footer").hide();
        $("#waitContainer").show();
        var fullUrl = location.protocol + '//' + location.hostname + (location.port ? ':' + location.port : '') + url;

        // height and width are percentages of the size of the parent Office application, e.g., PowerPoint, Excel, Word, etc.
        Office.context.ui.displayDialogAsync(fullUrl,
                {height: 60, width: 30}, function (result) {
            console.log("Dialog has initialized. Wiring up events");
            _dlg = result.value;
            _dlg.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, processMessage);
        });
    }

    function showResult(result) {
        return Excel.run(function(context) {
            const sheet = context.workbook.worksheets.getActiveWorksheet();

            const data = [[
                result[0].name, 
                result[1].name,
                result[2].name]];

            const range = sheet.getRange("B5:D5");
            range.values = data;
            range.format.autofitColumns();

            return context.sync();
        });
    }

}());

// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.

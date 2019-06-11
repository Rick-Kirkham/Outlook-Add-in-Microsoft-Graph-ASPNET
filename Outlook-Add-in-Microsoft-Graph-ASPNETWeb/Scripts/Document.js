﻿// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.

"use strict";

Office.initialize = function () {
    $(document).ready(function () {
        app.initialize();

        $("#getOneDriveFilesButton").click(getFileNamesFromGraph);
        $("#logoutO365PopupButton").click(function () {
            window.location.href = "/azureadauth/logout";
        });        
    });
};

function getFileNamesFromGraph() {

    $("#instructionsContainer").hide();
    $("#waitContainer").show();

    $.ajax({
        url: "/files/onedrivefiles",
        type: "GET"
    })
        .done(function (result) {
            writeFileNamesToMessage(result);
        })
        .then(function () {
            $("#waitContainer").hide();
            $("#finishedContainer").show();
        })
        .fail(function (result) {
            throw("Cannot get data from MS Graph: " + result);
        });
}

function writeFileNamesToMessage(graphData) {
    Office.context.mailbox.item.body.getTypeAsync(
        function (result) {
            if (result.status === Office.AsyncResultStatus.Failed) {
                console.log(result.error.message);
            }
            else {
                // Successfully got the type of item body.
                if (result.value === Office.MailboxEnums.BodyType.Html) {

                    // Body is of type HTML.
                    var htmlContent = createHtmlContent(graphData);

                    Office.context.mailbox.item.body.setSelectedDataAsync(
                        htmlContent, { coercionType: Office.CoercionType.Html },
                        function (asyncResult) {
                            if (asyncResult.status ===
                                Office.AsyncResultStatus.Failed) {
                                console.log(asyncResult.error.message);
                            }
                            else {
                                console.log("Successfully set HTML data in item body.");
                            }
                        });
                }
                else {
                    // Body is of type text. 
                    var textContent = createTextContent(graphData);

                    Office.context.mailbox.item.body.setSelectedDataAsync(
                        textContent, { coercionType: Office.CoercionType.Text },
                        function (asyncResult) {
                            if (asyncResult.status ===
                                Office.AsyncResultStatus.Failed) {
                                console.log(asyncResult.error.message);
                            }
                            else {
                                console.log("Successfully set text data in item body.");
                            }
                        });
                }
            }
        });
}

function createHtmlContent(data) {

    var bodyContent = "<html><head></head><body>";

    for (var i = 0; i < data.length; i++) {
        bodyContent += "<p>" + data[i] + "</p>";
    }
    bodyContent += "</body></html >";

    return bodyContent;
}

function createTextContent(data) {

    var bodyContent = "";
    for (var i = 0; i < data.length; i++) {
        bodyContent += data[i] + "\n";
    }

    return bodyContent;
}
// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.

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
            writeFileNamesToWorksheet(result);
        })
        .then(function () {
            $("#waitContainer").hide();
            $("#finishedContainer").show();
        })
        .fail(function (result) {
            throw("Cannot get data from MS Graph: " + result);
        });
}


function writeFileNamesToWorksheet(result) {
    
     return Excel.run(function (context) {
        const sheet = context.workbook.worksheets.getActiveWorksheet();

         const data = [
             [result[0]],
             [result[1]],
             [result[2]]];

        const range = sheet.getRange("B5:B7");
        range.values = data;
        range.format.autofitColumns();

        return context.sync();
    });

}
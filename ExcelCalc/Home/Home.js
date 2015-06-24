/// <reference path="../App.js" />

(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            $('#get-data-from-selection').click(getDataFromSelection);
            $('#write-range').click(writeRange);
        });
    };

    // Reads data from current document selection and displays a notification
    function getDataFromSelection() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    app.showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    app.showNotification('Error:', result.error.message);
                }
            }
        );
    }

    function writeRange() {
        var rangeAddress = "A1:A1";

        var ctx = new Excel.ExcelClientContext();
        var range = ctx.workbook.worksheets.getActiveWorksheet().getRange(rangeAddress);

        range.getCell(0, 0).values = "Hello World!";

        ctx.executeAsync().then(function () {
            app.showNotification("Write to Range"+rangeAddress+"is Successful!");
        }, function (error) {
            app.showNotification("Error", JSON.stringify(error));
        });
    }


})();